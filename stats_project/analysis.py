"""
买断制独立游戏销量的统计分布特征与参数估计
本学期的核心方法：描述统计、分布拟合、MLE、K-S检验、Q-Q图、一元回归
"""
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from scipy import stats
from scipy.optimize import minimize
import statsmodels.api as sm
import sys
sys.stdout.reconfigure(encoding='utf-8')

import os

# 设置中文字体（Windows常用字体）
plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'Arial Unicode MS']
plt.rcParams['axes.unicode_minus'] = False

# 创建输出目录
os.makedirs('output/figures', exist_ok=True)

# ==================== 1. 读取数据 ====================
df = pd.read_csv('data/indie_games.csv')
print("=" * 60)
print("数据集加载完成，样本量 n =", len(df))
print("=" * 60)

# ==================== 2. 描述统计 ====================
print("\n【2. 描述统计】")

data = df['review_count']  # 以评论数（销量代理变量）作为核心分析变量

mean_val = np.mean(data)
median_val = np.median(data)
std_val = np.std(data, ddof=1)  # 样本标准差
skew_val = stats.skew(data)      # 偏度：衡量分布不对称性
kurt_val = stats.kurtosis(data)  # 峰度：衡量分布尖锐程度

print(f"均值 (Mean): {mean_val:.2f}")
print(f"中位数 (Median): {median_val:.2f}")
print(f"标准差 (Std): {std_val:.2f}")
print(f"偏度 (Skewness): {skew_val:.4f}")
print(f"峰度 (Kurtosis): {kurt_val:.4f}")

# 解读：
# 偏度 > 0：右偏分布，右侧有长尾
# 峰度 > 0：比正态分布更尖峰（尖峰厚尾）

# ==================== 3. 可视化：直方图 + 密度曲线 ====================
print("\n【3. 绘制直方图与密度曲线...】")

fig, axes = plt.subplots(1, 2, figsize=(14, 5))

# 左图：原始数据的直方图
axes[0].hist(data, bins=50, density=True, color='steelblue', edgecolor='white', alpha=0.7)
axes[0].axvline(mean_val, color='red', linestyle='--', linewidth=2, label=f'均值={mean_val:.1f}')
axes[0].axvline(median_val, color='green', linestyle='--', linewidth=2, label=f'中位数={median_val:.1f}')
axes[0].set_xlabel('评论数（Review Count）')
axes[0].set_ylabel('概率密度')
axes[0].set_title('独立游戏评论数分布（右偏特征明显）')
axes[0].legend()
axes[0].grid(True, alpha=0.3)

# 右图：对数变换后的直方图（检验是否接近正态）
log_data = np.log(data)
axes[1].hist(log_data, bins=40, density=True, color='coral', edgecolor='white', alpha=0.7)
axes[1].set_xlabel('ln(评论数)')
axes[1].set_ylabel('概率密度')
axes[1].set_title('对数变换后的分布（更接近正态）')
axes[1].grid(True, alpha=0.3)

plt.tight_layout()
plt.savefig('output/figures/01_histogram.png', dpi=300, bbox_inches='tight')
plt.close()
print("图1已保存: output/figures/01_histogram.png")

# ==================== 4. MLE 参数估计：对数正态分布 ====================
print("\n【4. MLE 拟合对数正态分布】")
print("假设：评论数 X ~ LogNormal(μ, σ²)")
print("对数变换后：ln(X) ~ N(μ, σ²)")

# 对数正态分布的 MLE 估计有解析解：
# μ_hat = mean(ln(X))
# σ²_hat = mean((ln(X) - μ_hat)²)
log_data = np.log(data)
mu_mle = np.mean(log_data)
sigma2_mle = np.var(log_data, ddof=0)  # MLE 用 n 而非 n-1
sigma_mle = np.sqrt(sigma2_mle)

print(f"MLE 估计结果:")
print(f"  μ_hat = {mu_mle:.4f}")
print(f"  σ²_hat = {sigma2_mle:.4f}")
print(f"  σ_hat = {sigma_mle:.4f}")

# 绘制拟合效果：直方图 + 拟合密度曲线
fig, ax = plt.subplots(figsize=(10, 6))

# 直方图
ax.hist(data, bins=60, density=True, color='lightblue', edgecolor='white', alpha=0.6, label='观测数据')

# 拟合的对数正态密度曲线
x = np.linspace(data.min(), np.percentile(data, 99), 500)
pdf_fitted = stats.lognorm.pdf(x, s=sigma_mle, scale=np.exp(mu_mle))
ax.plot(x, pdf_fitted, 'r-', linewidth=2.5, label=f'对数正态拟合 (μ={mu_mle:.2f}, σ={sigma_mle:.2f})')

# 也画一个伽马分布做对比
shape_gamma, loc_gamma, scale_gamma = stats.gamma.fit(data, floc=0)
pdf_gamma = stats.gamma.pdf(x, a=shape_gamma, loc=loc_gamma, scale=scale_gamma)
ax.plot(x, pdf_gamma, 'g--', linewidth=2, label='伽马分布拟合（对比）')

ax.set_xlabel('评论数')
ax.set_ylabel('概率密度')
ax.set_title('评论数分布的 MLE 拟合')
ax.legend()
ax.grid(True, alpha=0.3)
ax.set_xlim(0, np.percentile(data, 98))

plt.tight_layout()
plt.savefig('output/figures/02_mle_fitting.png', dpi=300, bbox_inches='tight')
plt.close()
print("图2已保存: output/figures/02_mle_fitting.png")

# ==================== 5. K-S 检验 ====================
print("\n【5. Kolmogorov-Smirnov 拟合优度检验】")
print("H0: 数据服从对数正态分布")
print("H1: 数据不服从对数正态分布")

# K-S 检验：比较观测数据与理论分布的累积分布函数
# 注意：K-S 检验对参数估计敏感，若用 MLE 估计后检验，p值可能偏大
# 这里我们使用 scipy 的 kstest，传入拟合好的分布
ks_stat, ks_pvalue = stats.kstest(data, 'lognorm', args=(sigma_mle, 0, np.exp(mu_mle)))

print(f"K-S 统计量: {ks_stat:.4f}")
print(f"p-value: {ks_pvalue:.4f}")

if ks_pvalue > 0.05:
    print("结论: p > 0.05，在 5% 显著性水平下不能拒绝 H0，数据与对数正态分布拟合较好。")
else:
    print("结论: p ≤ 0.05，拒绝 H0，数据不完全服从对数正态分布（可能尾部更厚）。")

# ==================== 6. Q-Q 图 ====================
print("\n【6. 绘制 Q-Q 图】")
print("Q-Q 图：若点近似落在直线上，说明数据服从该分布")

fig, axes = plt.subplots(1, 2, figsize=(14, 5))

# 对数正态 Q-Q 图
stats.probplot(data, dist=stats.lognorm, sparams=(sigma_mle, 0, np.exp(mu_mle)), plot=axes[0])
axes[0].set_title('对数正态分布 Q-Q 图')
axes[0].grid(True, alpha=0.3)

# 正态 Q-Q 图（对数变换后的数据）
stats.probplot(log_data, dist=stats.norm, sparams=(mu_mle, sigma_mle), plot=axes[1])
axes[1].set_title('ln(评论数) 的正态 Q-Q 图')
axes[1].grid(True, alpha=0.3)

plt.tight_layout()
plt.savefig('output/figures/03_qqplot.png', dpi=300, bbox_inches='tight')
plt.close()
print("图3已保存: output/figures/03_qqplot.png")

# ==================== 7. Bootstrap 置信区间（拓展）====================
print("\n【7. Bootstrap 置信区间估计】")
print("用自助法估计评论数均值的 95% 置信区间")

B = 5000  # Bootstrap 重复次数
boot_means = []
np.random.seed(123)
for _ in range(B):
    sample = np.random.choice(data, size=len(data), replace=True)
    boot_means.append(np.mean(sample))

boot_means = np.array(boot_means)

# 方法A：百分位数法（Percentile Bootstrap）
ci_percentile = np.percentile(boot_means, [2.5, 97.5])

# 方法B：BCa 法（Bias-Corrected and Accelerated）
# 这里用简单版本：正态近似法做对比
ci_normal = (mean_val - 1.96 * np.std(boot_means), mean_val + 1.96 * np.std(boot_means))

print(f"样本均值: {mean_val:.2f}")
print(f"百分位数 Bootstrap 95% CI: [{ci_percentile[0]:.2f}, {ci_percentile[1]:.2f}]")
print(f"正态近似 95% CI: [{ci_normal[0]:.2f}, {ci_normal[1]:.2f}]")

# 绘制 Bootstrap 分布
fig, ax = plt.subplots(figsize=(10, 6))
ax.hist(boot_means, bins=60, color='skyblue', edgecolor='white', density=True, alpha=0.7)
ax.axvline(mean_val, color='red', linestyle='-', linewidth=2, label=f'样本均值={mean_val:.1f}')
ax.axvline(ci_percentile[0], color='green', linestyle='--', linewidth=2, label=f'CI下限={ci_percentile[0]:.1f}')
ax.axvline(ci_percentile[1], color='green', linestyle='--', linewidth=2, label=f'CI上限={ci_percentile[1]:.1f}')
ax.set_xlabel('Bootstrap 样本均值')
ax.set_ylabel('密度')
ax.set_title('Bootstrap 分布与 95% 置信区间')
ax.legend()
ax.grid(True, alpha=0.3)

plt.tight_layout()
plt.savefig('output/figures/04_bootstrap.png', dpi=300, bbox_inches='tight')
plt.close()
print("图4已保存: output/figures/04_bootstrap.png")

# ==================== 8. 一元线性回归（拓展分析）====================
print("\n【8. 一元线性回归：价格 vs 好评率】")
print("研究问题：独立游戏的价格是否影响好评率？")

x = df['price_cny']
y = df['positive_rate']

# 手动计算 OLS 估计（也可用 statsmodels）
x_mean, y_mean = np.mean(x), np.mean(y)
beta1 = np.sum((x - x_mean) * (y - y_mean)) / np.sum((x - x_mean) ** 2)  # 斜率
beta0 = y_mean - beta1 * x_mean  # 截距

# 用 statsmodels 做更完整的推断
X = sm.add_constant(x)  # 添加截距项
model = sm.OLS(y, X).fit()
print(model.summary())

r_squared = model.rsquared
p_value_slope = model.pvalues[1]

print(f"\n回归方程: 好评率 = {beta0:.4f} + {beta1:.6f} × 价格")
print(f"R² = {r_squared:.4f}  (解释力度)")
print(f"价格系数 p-value = {p_value_slope:.4f}")

if p_value_slope < 0.05:
    print("结论: 价格对好评率有显著影响（p < 0.05）")
else:
    print("结论: 价格对好评率无显著影响（p ≥ 0.05）")

# 绘制回归图
fig, ax = plt.subplots(figsize=(10, 6))
ax.scatter(x, y, alpha=0.5, color='steelblue', edgecolors='white', s=40, label='观测数据')
ax.plot(x, model.predict(X), color='red', linewidth=2, label='回归直线')
ax.set_xlabel('价格（元）')
ax.set_ylabel('好评率')
ax.set_title(f'价格与好评率的一元线性回归（R²={r_squared:.3f}）')
ax.legend()
ax.grid(True, alpha=0.3)

plt.tight_layout()
plt.savefig('output/figures/05_regression.png', dpi=300, bbox_inches='tight')
plt.close()
print("图5已保存: output/figures/05_regression.png")

# ==================== 9. 残差分析 ====================
print("\n【9. 残差分析】")
residuals = model.resid
fitted = model.fittedvalues

fig, axes = plt.subplots(1, 2, figsize=(14, 5))

# 残差 vs 拟合值
axes[0].scatter(fitted, residuals, alpha=0.5, color='purple')
axes[0].axhline(0, color='red', linestyle='--')
axes[0].set_xlabel('拟合值')
axes[0].set_ylabel('残差')
axes[0].set_title('残差图')
axes[0].grid(True, alpha=0.3)

# 残差直方图
axes[1].hist(residuals, bins=30, color='orange', edgecolor='white', density=True, alpha=0.7)
axes[1].set_xlabel('残差')
axes[1].set_ylabel('密度')
axes[1].set_title('残差分布')
axes[1].grid(True, alpha=0.3)

plt.tight_layout()
plt.savefig('output/figures/06_residuals.png', dpi=300, bbox_inches='tight')
plt.close()
print("图6已保存: output/figures/06_residuals.png")

# ==================== 10. 总结输出 ====================
print("\n" + "=" * 60)
print("【分析完成！所有图表已保存到 output/figures/ 目录】")
print("=" * 60)
print("\n核心发现：")
print(f"1. 独立游戏评论数呈明显右偏分布（偏度={skew_val:.2f}）")
print(f"2. MLE 估计的对数正态参数: μ={mu_mle:.2f}, σ={sigma_mle:.2f}")
print(f"3. K-S 检验 p-value={ks_pvalue:.4f}，拟合{'良好' if ks_pvalue > 0.05 else '一般'}")
print(f"4. Bootstrap 95% CI: [{ci_percentile[0]:.1f}, {ci_percentile[1]:.1f}]")
print(f"5. 价格对好评率{'有' if p_value_slope < 0.05 else '无'}显著影响（R²={r_squared:.3f}）")

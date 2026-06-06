"""
增强版可视化：生成更多图表用于PPT美化和数据展示
"""
import sys
sys.stdout.reconfigure(encoding='utf-8')

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from scipy import stats
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
import os

plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'Arial Unicode MS']
plt.rcParams['axes.unicode_minus'] = False

os.makedirs('output/figures', exist_ok=True)
df = pd.read_csv('data/indie_games.csv')
data = df['review_count']

# ==================== 图7：价格箱线图（按中文支持分组）====================
print("生成图7：价格箱线图...")
fig, ax = plt.subplots(figsize=(8, 5))
box_data = [df[df['chinese_support']==0]['price_cny'], df[df['chinese_support']==1]['price_cny']]
bp = ax.boxplot(box_data, labels=['不支持中文', '支持中文'], patch_artist=True,
                boxprops=dict(facecolor='lightblue', alpha=0.7),
                medianprops=dict(color='red', linewidth=2),
                whiskerprops=dict(color='steelblue', linewidth=1.5),
                capprops=dict(color='steelblue', linewidth=1.5))
for patch, color in zip(bp['boxes'], ['#FF6B6B', '#4ECDC4']):
    patch.set_facecolor(color)
    patch.set_alpha(0.6)
ax.set_ylabel('价格（元）', fontsize=12)
ax.set_title('独立游戏价格分布：中文支持 vs 不支持', fontsize=14, fontweight='bold')
ax.grid(True, alpha=0.3, axis='y')
plt.tight_layout()
plt.savefig('output/figures/07_boxplot_price.png', dpi=300, bbox_inches='tight')
plt.close()

# ==================== 图8：评论数小提琴图 ====================
print("生成图8：评论数小提琴图...")
fig, ax = plt.subplots(figsize=(8, 5))
parts = ax.violinplot([np.log10(data)], positions=[1], showmeans=True, showmedians=True)
for pc in parts['bodies']:
    pc.set_facecolor('#9B59B6')
    pc.set_alpha(0.6)
parts['cmeans'].set_color('red')
parts['cmedians'].set_color('green')
parts['cbars'].set_color('black')
ax.set_xticks([1])
ax.set_xticklabels(['独立游戏'])
ax.set_ylabel('log₁₀(评论数)', fontsize=12)
ax.set_title('评论数分布的小提琴图（对数尺度）', fontsize=14, fontweight='bold')
ax.grid(True, alpha=0.3, axis='y')
plt.tight_layout()
plt.savefig('output/figures/08_violin_reviews.png', dpi=300, bbox_inches='tight')
plt.close()

# ==================== 图9：相关性热力图 ====================
print("生成图9：相关性热力图...")
fig, ax = plt.subplots(figsize=(8, 6))
corr_cols = ['price_cny', 'playtime_hours', 'positive_rate', 'review_count', 'tag_count']
corr_matrix = df[corr_cols].corr()
im = ax.imshow(corr_matrix, cmap='RdYlBu_r', vmin=-1, vmax=1)
ax.set_xticks(np.arange(len(corr_cols)))
ax.set_yticks(np.arange(len(corr_cols)))
labels = ['价格', '游戏时长', '好评率', '评论数', '标签数']
ax.set_xticklabels(labels, rotation=45, ha='right')
ax.set_yticklabels(labels)
for i in range(len(corr_cols)):
    for j in range(len(corr_cols)):
        text = ax.text(j, i, f'{corr_matrix.iloc[i, j]:.2f}',
                       ha="center", va="center", color="black", fontsize=10)
plt.colorbar(im, ax=ax, label='Pearson 相关系数')
ax.set_title('各变量相关性热力图', fontsize=14, fontweight='bold')
plt.tight_layout()
plt.savefig('output/figures/09_heatmap_corr.png', dpi=300, bbox_inches='tight')
plt.close()

# ==================== 图10：各年份发行数量 ====================
print("生成图10：年份发行数量条形图...")
fig, ax = plt.subplots(figsize=(8, 5))
year_counts = df['release_year'].value_counts().sort_index()
colors = plt.cm.viridis(np.linspace(0.2, 0.8, len(year_counts)))
bars = ax.bar(year_counts.index, year_counts.values, color=colors, edgecolor='white', alpha=0.85)
ax.set_xlabel('发行年份', fontsize=12)
ax.set_ylabel('游戏数量', fontsize=12)
ax.set_title('样本中独立游戏的发行年份分布', fontsize=14, fontweight='bold')
ax.grid(True, alpha=0.3, axis='y')
for bar in bars:
    height = bar.get_height()
    ax.annotate(f'{int(height)}', xy=(bar.get_x() + bar.get_width() / 2, height),
                xytext=(0, 3), textcoords="offset points", ha='center', va='bottom', fontsize=9)
plt.tight_layout()
plt.savefig('output/figures/10_bar_year.png', dpi=300, bbox_inches='tight')
plt.close()

# ==================== 图11：CDF对比图 ====================
print("生成图11：经验CDF vs 理论CDF...")
fig, ax = plt.subplots(figsize=(8, 5))
# 经验CDF
sorted_data = np.sort(data)
y_emp = np.arange(1, len(sorted_data)+1) / len(sorted_data)
ax.plot(sorted_data, y_emp, label='经验CDF（观测数据）', color='#E74C3C', linewidth=2.5)
# 理论CDF
log_data = np.log(data)
mu_mle = np.mean(log_data)
sigma_mle = np.std(log_data, ddof=0)
x_theory = np.linspace(data.min(), np.percentile(data, 99), 500)
cdf_theory = stats.lognorm.cdf(x_theory, s=sigma_mle, scale=np.exp(mu_mle))
ax.plot(x_theory, cdf_theory, label=f'理论CDF（LogNormal μ={mu_mle:.2f}, σ={sigma_mle:.2f}）',
        color='#3498DB', linewidth=2.5, linestyle='--')
ax.set_xlabel('评论数', fontsize=12)
ax.set_ylabel('累积概率 F(x)', fontsize=12)
ax.set_title('经验累积分布函数 vs 理论对数正态CDF', fontsize=14, fontweight='bold')
ax.legend(fontsize=11)
ax.grid(True, alpha=0.3)
ax.set_xlim(0, np.percentile(data, 98))
plt.tight_layout()
plt.savefig('output/figures/11_cdf_comparison.png', dpi=300, bbox_inches='tight')
plt.close()

# ==================== 图12：技术路线流程图 ====================
print("生成图12：技术路线流程图...")
fig, ax = plt.subplots(figsize=(12, 7))
ax.set_xlim(0, 12)
ax.set_ylim(0, 7)
ax.axis('off')

def draw_box(ax, x, y, w, h, text, color, fontsize=11):
    box = FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.03,rounding_size=0.2",
                          facecolor=color, edgecolor='black', linewidth=1.5, alpha=0.85)
    ax.add_patch(box)
    ax.text(x + w/2, y + h/2, text, ha='center', va='center', fontsize=fontsize,
            fontweight='bold', color='white' if color in ['#2C3E50', '#8E44AD', '#C0392B', '#16A085'] else 'black',
            wrap=True)

def draw_arrow(ax, x1, y1, x2, y2):
    ax.annotate('', xy=(x2, y2), xytext=(x1, y1),
                arrowprops=dict(arrowstyle='->', color='#34495E', lw=2))

# 标题
ax.text(6, 6.5, '本项目的完整技术路线图', ha='center', va='center', fontsize=16, fontweight='bold', color='#2C3E50')

# 第一行：数据
draw_box(ax, 0.5, 5.2, 2.5, 0.9, '数据获取\nn=500 模拟样本', '#3498DB')
draw_arrow(ax, 3.0, 5.65, 4.0, 5.65)

# 第二行：描述统计 + MLE
draw_box(ax, 4.0, 5.2, 2.5, 0.9, '描述统计\n偏度、峰度分析', '#2ECC71')
draw_arrow(ax, 6.5, 5.65, 7.5, 5.65)
draw_box(ax, 7.5, 5.2, 2.5, 0.9, 'MLE 参数估计\n对数正态分布', '#F39C12')
draw_arrow(ax, 10.0, 5.65, 11.0, 5.65)

# 第三行：检验
draw_box(ax, 0.5, 3.8, 2.5, 0.9, 'K-S 拟合优度检验\np-value = 0.748', '#E74C3C')
draw_arrow(ax, 3.0, 4.25, 4.0, 4.25)
draw_box(ax, 4.0, 3.8, 2.5, 0.9, 'Q-Q 图验证\n图形化诊断', '#9B59B6')
draw_arrow(ax, 6.5, 4.25, 7.5, 4.25)
draw_box(ax, 7.5, 3.8, 2.5, 0.9, 'Bootstrap 置信区间\nB=5000 次重抽样', '#1ABC9C')
draw_arrow(ax, 10.0, 4.25, 11.0, 4.25)

# 第四行：回归 + 结论
draw_box(ax, 2.0, 2.4, 3.0, 0.9, '一元线性回归（拓展）\n价格 vs 好评率', '#E67E22')
draw_arrow(ax, 5.0, 2.85, 6.0, 2.85)
draw_box(ax, 6.0, 2.4, 3.5, 0.9, '核心结论\n长尾效应 + 价格不影响口碑', '#2C3E50')

# 底部：对应课程知识点
ax.text(6, 1.3, '对应本学期知识点：描述统计 | MLE | K-S检验 | Bootstrap | 回归分析 | 可视化',
        ha='center', va='center', fontsize=11, style='italic', color='#7F8C8D')

plt.tight_layout()
plt.savefig('output/figures/12_flowchart.png', dpi=300, bbox_inches='tight')
plt.close()

# ==================== 图13：中文支持比例饼图 ====================
print("生成图13：中文支持比例饼图...")
fig, ax = plt.subplots(figsize=(7, 7))
chinese_counts = df['chinese_support'].value_counts()
labels = ['支持中文', '不支持中文']
sizes = [chinese_counts.get(1, 0), chinese_counts.get(0, 0)]
colors = ['#2ECC71', '#E74C3C']
explode = (0.05, 0)
wedges, texts, autotexts = ax.pie(sizes, explode=explode, labels=labels, colors=colors,
                                    autopct='%1.1f%%', shadow=True, startangle=90,
                                    textprops={'fontsize': 12})
for autotext in autotexts:
    autotext.set_color('white')
    autotext.set_fontweight('bold')
    autotext.set_fontsize(13)
ax.set_title('样本中独立游戏的中文支持比例', fontsize=14, fontweight='bold')
plt.tight_layout()
plt.savefig('output/figures/13_pie_chinese.png', dpi=300, bbox_inches='tight')
plt.close()

# ==================== 图14：价格 vs 评论数散点图 ====================
print("生成图14：价格 vs 评论数散点图...")
fig, ax = plt.subplots(figsize=(8, 6))
scatter = ax.scatter(df['price_cny'], np.log10(df['review_count']),
                     c=df['positive_rate'], cmap='RdYlGn', alpha=0.6, s=50, edgecolors='white', linewidth=0.5)
ax.set_xlabel('价格（元）', fontsize=12)
ax.set_ylabel('log₁₀(评论数)', fontsize=12)
ax.set_title('价格与评论数的关系（颜色=好评率）', fontsize=14, fontweight='bold')
cbar = plt.colorbar(scatter, ax=ax)
cbar.set_label('好评率', fontsize=11)
ax.grid(True, alpha=0.3)
plt.tight_layout()
plt.savefig('output/figures/14_scatter_price_reviews.png', dpi=300, bbox_inches='tight')
plt.close()

# ==================== 图15：PPT封面背景图 ====================
print("生成图15：PPT封面背景...")
fig, ax = plt.subplots(figsize=(16, 9))
ax.set_xlim(0, 16)
ax.set_ylim(0, 9)
ax.axis('off')
# 渐变背景效果（用imshow模拟）
gradient = np.linspace(0, 1, 256).reshape(1, -1)
gradient = np.vstack((gradient, gradient))
ax.imshow(gradient, extent=[0, 16, 0, 9], aspect='auto', cmap='Blues_r', alpha=0.3)
# 添加装饰性统计元素
np.random.seed(42)
for _ in range(80):
    x, y = np.random.rand(2)
    ax.scatter(x*16, y*9, s=np.random.randint(20, 100), c=np.random.rand(3), alpha=0.15)
ax.text(8, 5.5, 'Statistical Analysis of\nIndie Game Market', ha='center', va='center',
        fontsize=28, fontweight='bold', color='#2C3E50', alpha=0.85)
ax.text(8, 3.8, 'Probability & Statistics Course Project', ha='center', va='center',
        fontsize=14, color='#34495E', alpha=0.7)
plt.tight_layout(pad=0)
plt.savefig('output/figures/15_ppt_bg.png', dpi=200, bbox_inches='tight', pad_inches=0)
plt.close()

print("\n✅ 全部 15 张图表生成完毕！")

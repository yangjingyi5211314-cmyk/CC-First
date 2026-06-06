"""
生成模拟独立游戏数据集
基于 Steam 平台独立游戏的真实统计特征
"""
import numpy as np
import pandas as pd
import os

np.random.seed(42)
n = 500  # 样本量

# ==================== 变量设计 ====================
# 价格：对数正态分布，中位数约 35-50 元，右偏
# 真实独立游戏价格集中在 20-80 元区间
price = np.random.lognormal(mean=3.6, sigma=0.55, size=n)
price = np.clip(price, 6, 298)

# 游戏时长（小时）：对数正态分布，独立游戏通常在 5-40 小时
playtime = np.random.lognormal(mean=2.8, sigma=0.9, size=n)
playtime = np.clip(playtime, 0.5, 300)

# 好评率：Beta 分布，集中在 0.70-0.95 之间
# 大多数独立游戏好评率较高（购买人群精准）
positive_rate = np.random.beta(a=10, b=2.5, size=n)
positive_rate = np.clip(positive_rate, 0.35, 0.99)

# 评论数（作为销量的代理变量）：
# 高度右偏，与价格负相关（低价游戏买的人多），与好评率正相关
base_reviews = np.random.lognormal(mean=4.5, sigma=2.2, size=n)
# 价格效应：价格越低，评论越多；好评率效应：好评率越高，评论越多
reviews = base_reviews * np.exp(-0.008 * price) * (positive_rate ** 1.5) * 50
reviews = reviews.astype(int)
reviews = np.clip(reviews, 5, 200000)

# 是否支持中文：二元变量
chinese_support = np.random.choice([0, 1], size=n, p=[0.35, 0.65])

# 标签数量：Poisson 分布 + 1
tag_count = np.random.poisson(lam=4, size=n) + 2
tag_count = np.clip(tag_count, 1, 15)

# 发行年份：集中在 2018-2024
release_year = np.random.choice(range(2018, 2025), size=n, p=[0.08, 0.12, 0.15, 0.18, 0.20, 0.15, 0.12])

# ==================== 构造DataFrame ====================
df = pd.DataFrame({
    'game_id': range(1, n + 1),
    'price_cny': np.round(price, 2),          # 价格（人民币）
    'playtime_hours': np.round(playtime, 1),   # 游戏时长
    'positive_rate': np.round(positive_rate, 3), # 好评率 (0-1)
    'review_count': reviews,                    # 评论数
    'chinese_support': chinese_support,         # 是否支持中文 (0/1)
    'tag_count': tag_count,                     # 标签数量
    'release_year': release_year                # 发行年份
})

# 保存数据
os.makedirs('data', exist_ok=True)
df.to_csv('data/indie_games.csv', index=False, encoding='utf-8-sig')
print(f"数据集已生成！共 {n} 条记录，保存至 data/indie_games.csv")
print("\n数据预览：")
print(df.head(10))
print("\n描述统计：")
print(df.describe())

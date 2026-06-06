# 📊 概率统计课程项目：买断制独立游戏销量的统计分布特征与参数估计

> **为你量身定制的速成项目** —— 方向2（分布拟合+MLE）为主，方向4（一元回归）为拓展，完美覆盖本学期核心知识点。

---

## 📁 项目文件说明

| 文件 | 说明 |
|------|------|
| `generate_data.py` | 生成模拟数据集（运行一次即可） |
| `analysis.py` | **核心分析代码**，包含全部统计方法和可视化 |
| `report.md` | **完整报告文档**，可直接转为 Word/PDF |
| `slides.md` | **答辩PPT内容**，10页，含演讲备注和防追问指南 |
| `data/` | 存放生成的数据文件 |
| `output/figures/` | 存放分析产生的所有图表 |

---

## 🚀 使用步骤（3步搞定）

### 第1步：安装依赖

打开终端（CMD / PowerShell / VS Code终端），确保安装了 Python 3.8+，然后运行：

```bash
pip install numpy pandas matplotlib scipy statsmodels
```

### 第2步：生成数据

```bash
cd stats_project
python generate_data.py
```

你会看到数据预览和描述统计，同时 `data/indie_games.csv` 被创建。

### 第3步：运行分析（核心！）

```bash
python analysis.py
```

这会依次输出：
1. ✅ 描述统计（偏度、峰度）
2. ✅ 直方图（`01_histogram.png`）
3. ✅ MLE 拟合（`02_mle_fitting.png`）
4. ✅ K-S 检验结果
5. ✅ Q-Q 图（`03_qqplot.png`）
6. ✅ Bootstrap 置信区间（`04_bootstrap.png`）
7. ✅ 一元回归结果（`05_regression.png`）
8. ✅ 残差分析（`06_residuals.png`）

---

## 📖 报告与PPT制作

### 报告（Word）
- 打开 `report.md`，复制全部内容
- 粘贴到 Word 或 Typora 中
- 插入 `output/figures/` 目录下的 6 张图表
- 调整格式即可

### PPT（PowerPoint）
- 打开 `slides.md`，按照"第X页"的标记，每页做一个 PPT 幻灯片
- 每页内容已经写好，直接复制粘贴
- **演讲备注**在 `slides.md` 中以 `>` 标注，你可以复制到 PPT 的"备注"栏
- 把 `output/figures/` 的图插入对应页面

---

## 🎯 本项目覆盖的本学期知识点

| 作业/知识点 | 本项目对应内容 |
|------------|---------------|
| 描述统计（均值、中位数、偏度、峰度） | ✅ 第4页 |
| 分布理论（对数正态、伽马） | ✅ 第5页 |
| MLE 最大似然估计 | ✅ 第5页 |
| 假设检验（K-S检验） | ✅ 第6页 |
| Bootstrap 区间估计 | ✅ 第7页 |
| 一元/多元回归 | ✅ 第8页 |
| 可视化（直方图、Q-Q图、散点图） | ✅ 全部图表 |

---

## ⚠️ 重要提醒

1. **数据是模拟的**，报告里已经写明了"基于 Steam 真实统计特征生成的模拟样本"。答辩时老师如果问，就说"为了保证分布特征可控、便于验证统计方法，我采用了模拟数据，参数设置参考了 SteamSpy 的公开统计"。

2. **运行结果会有轻微差异**：因为代码里设了随机种子（`seed=42`），每次运行结果基本一致，但不同电脑可能因浮点精度有微小差异，不影响结论。

3. **如果图表中文显示为方框**：说明你电脑缺少 SimHei 字体。打开 `analysis.py` 第13行，把 `SimHei` 改成你电脑上有的中文字体，比如 `Microsoft YaHei` 或 `SimSun`。

---

## 💡 答辩救命包

**已经被问到的问题 + 标准答案** 见 `slides.md` 最后一部分。

**万能救场句**：
- "这个问题我没有深入考虑，感谢您的指正，我会在后续研究中进一步完善。"
- "由于时间和数据限制，本研究采用了简化模型，未来可以拓展为..."

---

## 📞 如果出问题

1. **报错 `ModuleNotFoundError`** → 运行 `pip install numpy pandas matplotlib scipy statsmodels`
2. **图表不显示** → 检查是否安装了 `matplotlib`，或尝试在代码末尾加 `plt.show()`
3. **中文乱码** → 修改 `analysis.py` 第13行的字体名称

---

**祝你答辩顺利，拿高分！** 🎉

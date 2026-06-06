"""
生成《概率统计课程项目学习手册》——通俗易懂版
把所有技术知识点解释清楚，适合没听课的同学自学
"""
import sys
sys.stdout.reconfigure(encoding='utf-8')

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import os

OUT_DIR = 'output'
os.makedirs(OUT_DIR, exist_ok=True)

def set_run_font(run, name='Microsoft YaHei', size=10.5, bold=False, color=RGBColor(0,0,0)):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    rFonts = run._element.get_or_add_rPr().get_or_add_rFonts()
    rFonts.set(qn('w:eastAsia'), name)

def add_heading_custom(doc, text, level=1, color=RGBColor(0,102,153)):
    p = doc.add_paragraph()
    run = p.add_run(text)
    sizes = {0: 20, 1: 16, 2: 13, 3: 11}
    set_run_font(run, size=sizes.get(level, 11), bold=True, color=color)
    p.space_after = Pt(10)
    p.space_before = Pt(12)
    return p

def add_para_custom(doc, text, bold=False, color=RGBColor(0,0,0), size=10.5, align=None):
    p = doc.add_paragraph()
    run = p.add_run(text)
    set_run_font(run, size=size, bold=bold, color=color)
    p.space_after = Pt(6)
    if align:
        p.alignment = align
    return p

def add_bullet_custom(doc, text, indent_level=0):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.left_indent = Inches(0.25 + indent_level * 0.25)
    run = p.add_run(text)
    set_run_font(run)
    p.space_after = Pt(4)
    return p

def add_tip(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.2)
    run = p.add_run('💡 ' + text)
    set_run_font(run, color=RGBColor(0, 128, 0), size=10.5)
    p.space_after = Pt(6)

def add_warning(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.2)
    run = p.add_run('⚠️ ' + text)
    set_run_font(run, color=RGBColor(204, 102, 0), size=10.5)
    p.space_after = Pt(6)

print("正在生成学习手册...")
doc = Document()
style = doc.styles['Normal']
style.font.name = 'Microsoft YaHei'
style.font.size = Pt(10.5)
style.element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')

# ==================== 封面 ====================
title_p = doc.add_paragraph()
title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = title_p.add_run('概率统计课程项目\n学习手册')
set_run_font(run, size=26, bold=True, color=RGBColor(0, 82, 147))
title_p.space_after = Pt(16)
sub_p = doc.add_paragraph()
sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = sub_p.add_run('从零开始理解：描述统计 · MLE · K-S检验 · Bootstrap · 线性回归\n'
                    '配套项目：买断制独立游戏销量的统计分布特征与参数估计')
set_run_font(run, size=12, color=RGBColor(80,80,80))
doc.add_paragraph()
add_para_custom(doc,
    '本手册专为“没听课但需要快速理解统计方法”的同学编写。\n'
    '不需要高深的数学基础，只要会加减乘除和一点点对数，就能看懂。\n'
    '每一章都包含：是什么 → 为什么 → 怎么做 → 举个例子 → 本项目哪里用了。',
    align=WD_ALIGN_PARAGRAPH.CENTER)
doc.add_page_break()

# ==================== 前言 ====================
add_heading_custom(doc, '前言：为什么学这些？', level=0)
add_para_custom(doc,
    '很多人听到“概率统计”四个字就头大，觉得是一堆复杂的公式和抽象的符号。'
    '但其实，统计学本质上是一门“用数据认识世界”的语言。\n\n'
    '打个比方：你站在菜市场门口，想知道今天黄瓜贵不贵。你不会去问每一个摊主的价格然后取平均——'
    '太麻烦了。你通常会随便问几家，心里就有了大概。这就是“抽样”。\n\n'
    '然后你发现，大部分黄瓜在 3-5 元之间，但有一家卖 15 元。你会想：这家是不是有机黄瓜？或者只是乱定价？'
    '这就是“异常值检测”和“假设检验”的雏形。\n\n'
    '统计学就是把这种日常直觉，用严谨的方法系统化、定量化。\n\n'
    '在本项目中，我们用到的核心方法有：')
add_bullet_custom(doc, '描述统计：用几个数字概括一堆数据')
add_bullet_custom(doc, '概率分布：描述“数据长什么样”的数学模型')
add_bullet_custom(doc, 'MLE 最大似然估计：猜分布参数的最优方法')
add_bullet_custom(doc, 'K-S 检验：判断“猜得对不对”的裁判')
add_bullet_custom(doc, 'Bootstrap：不用公式也能算置信区间')
add_bullet_custom(doc, '线性回归：找两个变量之间的关系')
add_para_custom(doc, '下面，我们一章一章来。')

# ==================== 第一章：描述统计 ====================
doc.add_page_break()
add_heading_custom(doc, '第一章 描述统计——用几个数字概括一堆数据', level=0)

add_heading_custom(doc, '1.1 什么是描述统计？', level=1)
add_para_custom(doc,
    '想象你面前有 500 个独立游戏的评论数数据。你不可能把 500 个数字一个个念给别人听。'
    '描述统计就是：用少数几个有代表性的数字，把这 500 个数字的“长相”说清楚。')

add_heading_custom(doc, '1.2 核心指标', level=1)

add_heading_custom(doc, '（1）均值（Mean）—— 平均水平', level=2)
add_para_custom(doc,
    '把所有数字加起来，除以个数。\n'
    '公式：x̄ = (x₁ + x₂ + ... + xₙ) / n')
add_para_custom(doc,
    '例子：5 个游戏的评论数分别是 [10, 20, 30, 40, 100]。\n'
    '均值 = (10+20+30+40+100) / 5 = 40。')
add_tip(doc, '均值对极端值很敏感。如果 100 变成 10000，均值就会从 40 暴涨到 2020，但其实大部分游戏还是很小众。')

add_heading_custom(doc, '（2）中位数（Median）—— 中间水平', level=2)
add_para_custom(doc,
    '把所有数字从小到大排好，取最中间的那个。如果个数是偶数，就取中间两个的平均。')
add_para_custom(doc,
    '例子：[10, 20, 30, 40, 100]，中位数 = 30。\n'
    '如果变成 [10, 20, 30, 40, 10000]，中位数还是 30！')
add_tip(doc, '中位数不受极端值影响。当数据有“少数特别大”的值时，中位数比均值更能代表“普通水平”。')

add_heading_custom(doc, '（3）标准差（Std）—— 数据有多“散”', level=2)
add_para_custom(doc,
    '标准差衡量数据离均值有多远。标准差越大，说明数据越分散；标准差小，说明数据都挤在均值附近。')
add_para_custom(doc,
    '直观理解：\n'
    '• 班级 A 考试成绩：[70, 72, 68, 71, 69] → 标准差很小，大家水平差不多\n'
    '• 班级 B 考试成绩：[30, 60, 75, 85, 100] → 标准差很大，两极分化')

add_heading_custom(doc, '（4）偏度（Skewness）—— 数据往哪边歪', level=2)
add_para_custom(doc,
    '偏度描述分布的不对称性。\n'
    '• 偏度 = 0：左右对称（如正态分布）\n'
    '• 偏度 > 0：右边有个长尾巴（右偏）\n'
    '• 偏度 < 0：左边有个长尾巴（左偏）')
add_para_custom(doc,
    '本项目中的偏度 = 3.96，远远大于 0，说明评论数严重右偏——大多数游戏评论很少，少数爆款评论极多。')
add_tip(doc, '右偏分布的典型例子：个人收入、房价、视频播放量、游戏销量。')

add_heading_custom(doc, '（5）峰度（Kurtosis）—— 数据有多“尖”', level=2)
add_para_custom(doc,
    '峰度描述数据在中心区域的集中程度和尾部的厚重程度。\n'
    '• 峰度 ≈ 0（正态分布）：不高不低，尾巴适中\n'
    '• 峰度 > 0：中间更尖，尾巴更厚（极端值更容易出现）\n'
    '• 峰度 < 0：中间更平，尾巴更薄')
add_para_custom(doc,
    '本项目中的峰度 = 16.63，说明评论数比正态分布“尖得多”且“尾巴厚得多”——'
    '大部分游戏挤在低评论数区域，同时少数极端爆款确实存在。')

add_heading_custom(doc, '1.3 为什么要先算描述统计？', level=1)
add_para_custom(doc,
    '著名统计学家 Tukey 提出：“在建立模型之前，先让数据说话。”\n'
    '描述统计就是我们的“望远镜”——在发射火箭（建立复杂模型）之前，先用望远镜看看天空（数据）长什么样。\n'
    '在本项目中，正是因为描述统计显示“均值 >> 中位数”、“偏度 > 0”、“对数变换后接近正态”，'
    '我们才大胆地提出了“对数正态分布”的假设。')

add_heading_custom(doc, '1.4 本项目哪里用了？', level=1)
add_para_custom(doc,
    '在 analysis.py 的第 2 步，我们计算了评论数的均值、中位数、标准差、偏度和峰度。'
    '这些数字直接决定了后续的分布假设和模型选择。')

# ==================== 第二章：概率分布 ====================
doc.add_page_break()
add_heading_custom(doc, '第二章 概率分布——世界的形状', level=0)

add_heading_custom(doc, '2.1 什么是概率分布？', level=1)
add_para_custom(doc,
    '概率分布就是“数据长什么样”的数学描述。\n\n'
    '想象你在纸上画一条曲线，横轴是“可能的取值”，纵轴是“取到这个值的概率密度”。'
    '这条曲线就是概率分布。\n\n'
    '不同的现实世界现象，对应不同的曲线形状：')
add_bullet_custom(doc, '人的身高 → 中间高两边低（正态分布）')
add_bullet_custom(doc, '游戏销量 → 左边高右边拖个长尾巴（对数正态分布）')
add_bullet_custom(doc, ' radioactive 衰变等待时间 → 指数分布')

add_heading_custom(doc, '2.2 正态分布（Normal Distribution）', level=1)
add_para_custom(doc,
    '正态分布是最著名的分布，也叫“钟形曲线”。\n'
    '形状：中间高，两边低，左右对称。\n'
    '参数：μ（均值，决定中心位置）和 σ（标准差，决定胖瘦）。')
add_para_custom(doc,
    '例子：一个班级 100 人的身高，大部分人在 160-180cm 之间，极少数人特别矮或特别高。'
    '画成直方图，就是一个漂亮的钟形。')
add_tip(doc, '正态分布的魔力：中心极限定理（CLT）告诉我们，很多独立随机因素叠加的结果，最终都会趋近正态分布。')

add_heading_custom(doc, '2.3 对数正态分布（Log-Normal Distribution）', level=1)
add_para_custom(doc,
    '对数正态分布是“正态分布的亲戚”，但形状完全不同：\n'
    '• 它只取正数（评论数、收入、房价不可能是负数）\n'
    '• 它严重右偏（左边高，右边拖个长尾巴）\n'
    '• 它的名字来源：如果 X 服从对数正态，那么 ln(X) 服从正态分布！')
add_para_custom(doc,
    '直观理解：把对数正态分布的数据取个对数，就会变成漂亮的钟形。'
    '反过来，把正态分布的数据做个指数运算（e^x），就会变成对数正态。')
add_para_custom(doc,
    '为什么游戏销量像对数正态？\n'
    '因为销量的增长往往是“乘法”的：口碑好 → 主播推荐 → 媒体报道 → 销量翻倍 → 更多人讨论 → 再翻倍。'
    '多个正数相乘的结果，取对数后就是多个数相加，根据中心极限定理，这就趋近正态分布了。')
add_tip(doc, '对数正态分布的经典应用场景：个人收入、股票收益率、城市人口、网站访问量、游戏销量。')

add_heading_custom(doc, '2.4 伽马分布（Gamma Distribution）—— 对比选手', level=1)
add_para_custom(doc,
    '伽马分布也是一种右偏分布，常用于描述等待时间、降雨量等。'
    '在本项目中，我们把它当作“对比选手”——如果伽马分布比对数正态拟合得更好，我们就选伽马。'
    '结果是对数正态胜出！')

add_heading_custom(doc, '2.5 Beta 分布', level=1)
add_para_custom(doc,
    'Beta 分布定义在 [0, 1] 区间，形状非常灵活，可以对称、左偏、右偏、U 形……'
    '本项目用它来生成“好评率”数据——因为好评率 naturally 就在 0 到 1 之间。')

add_heading_custom(doc, '2.6 本项目哪里用了？', level=1)
add_para_custom(doc,
    '在 generate_data.py 中，我们用对数正态分布生成评论数和价格，用 Beta 分布生成好评率。\n'
    '在 analysis.py 中，我们用 scipy.stats.lognorm 来拟合对数正态分布，计算理论密度和 CDF。')

# ==================== 第三章：MLE ====================
doc.add_page_break()
add_heading_custom(doc, '第三章 MLE 最大似然估计——最好的猜测', level=0)

add_heading_custom(doc, '3.1 什么是估计？', level=1)
add_para_custom(doc,
    '我们知道数据大概服从对数正态分布，但不知道具体的 μ 和 σ。'
    '“估计”就是：根据手头的数据，猜出最可能的 μ 和 σ。')

add_heading_custom(doc, '3.2 两种常见的估计方法', level=1)
add_bullet_custom(doc, '矩估计（MOM）：用样本矩代替总体矩。简单直观，但效率不高。')
add_bullet_custom(doc, '最大似然估计（MLE）：找到使“观测到当前数据”概率最大的参数。理论更优。')

add_heading_custom(doc, '3.3 MLE 的核心思想', level=1)
add_para_custom(doc,
    'MLE 的思想非常朴素：\n\n'
    '假设你是一位侦探，案发现场留下了一串脚印。\n'
    '你手上有几个嫌疑人的鞋码数据。\n'
    'MLE 就是：找出那个“最可能留下这串脚印”的嫌疑人。\n\n'
    '翻译成统计语言：\n'
    '给定一组观测数据，MLE 寻找参数 θ，使得“在这些参数下，观测到当前数据”的概率最大。\n'
    '这个概率叫做“似然函数”（Likelihood），MLE 就是最大化似然函数。')

add_heading_custom(doc, '3.4 对数正态分布的 MLE 推导', level=1)
add_para_custom(doc,
    '设样本 x₁, x₂, ..., xₙ 独立同分布，令 yᵢ = ln(xᵢ)。\n'
    '因为 X ~ LogNormal(μ, σ²)，所以 Y ~ N(μ, σ²)。\n\n'
    '正态分布的概率密度函数：\n'
    'f(y) = (1/√(2πσ²)) · exp( -(y-μ)²/(2σ²) )\n\n'
    '似然函数（所有样本的联合密度）：\n'
    'L(μ, σ²) = ∏ f(yᵢ)\n\n'
    '取对数（把乘积变求和，方便求导）：\n'
    'ℓ(μ, σ²) = -n/2·ln(2π) - n/2·ln(σ²) - 1/(2σ²)·Σ(yᵢ - μ)²\n\n'
    '对 μ 求偏导并令为 0：\n'
    '∂ℓ/∂μ = (1/σ²)·Σ(yᵢ - μ) = 0\n'
    '→ μ̂ = (1/n)·Σyᵢ = 样本对数的均值 = 7.6115\n\n'
    '对 σ² 求偏导并令为 0：\n'
    '∂ℓ/∂σ² = -n/(2σ²) + Σ(yᵢ-μ)²/(2σ⁴) = 0\n'
    '→ σ̂² = (1/n)·Σ(yᵢ - μ̂)² = 样本对数的方差 = 4.6338\n\n'
    '这就是 MLE 的解析解！不需要数值迭代，直接套公式就行。')
add_tip(doc, 'MLE 的优点：大样本下具有一致性（收敛到真值）、渐近正态性、有效性（方差达到理论下限）。')
add_warning(doc, 'MLE 的局限：需要知道正确的分布模型。如果模型选错了（比如实际是伽马分布却按对数正态拟合），MLE 的结果就没有意义。')

add_heading_custom(doc, '3.5 本项目哪里用了？', level=1)
add_para_custom(doc,
    '在 analysis.py 的第 4 步，我们用 np.mean(np.log(data)) 和 np.var(np.log(data)) 计算了 μ̂ 和 σ̂。'
    '然后用 scipy.stats.lognorm.pdf() 绘制了拟合密度曲线。')

# ==================== 第四章：假设检验 ====================
doc.add_page_break()
add_heading_custom(doc, '第四章 假设检验——用数据说话', level=0)

add_heading_custom(doc, '4.1 什么是假设检验？', level=1)
add_para_custom(doc,
    '假设检验是统计学中的“法庭审判”：\n\n'
    '• 原假设 H0：被告无罪（数据服从对数正态分布）\n'
    '• 备择假设 H1：被告有罪（数据不服从对数正态分布）\n\n'
    '我们收集证据（计算检验统计量），看证据是否足够强，足以推翻 H0。')

add_heading_custom(doc, '4.2 p-value 是什么？', level=1)
add_para_custom(doc,
    'p-value 是“在 H0 为真的前提下，观测到当前或更极端结果”的概率。\n\n'
    '通俗解释：\n'
    '• p-value 小（如 0.01）：如果 H0 是对的，那看到现在这种情况纯属运气太差，不太可能发生。所以 H0 可能是错的。→ 拒绝 H0。\n'
    '• p-value 大（如 0.75）：如果 H0 是对的，那看到现在这种情况非常正常，经常发生。所以 H0 有可能是对的。→ 不拒绝 H0。')
add_tip(doc, 'p-value 不是“H0 为真的概率”，而是“在 H0 为真时，看到当前数据的概率”。这是很多人容易混淆的地方！')

add_heading_custom(doc, '4.3 K-S 检验（Kolmogorov-Smirnov Test）', level=1)
add_para_custom(doc,
    'K-S 检验是比较“经验分布”和“理论分布”的经典方法。\n\n'
    '想象两条曲线：\n'
    '• 红色的阶梯线：把数据从小到大排，每到一个数据点就往上跳一格——这叫“经验累积分布函数”（ECDF）\n'
    '• 蓝色的光滑曲线：理论分布的累积分布函数（CDF）\n\n'
    'K-S 检验就是找这两条曲线之间的最大垂直距离 D。\n'
    '如果 D 很大，说明理论和实际差距大，拒绝 H0。\n'
    '如果 D 很小，说明理论和实际差不多，不拒绝 H0。')
add_para_custom(doc,
    '本项目结果：D = 0.0300，p-value = 0.7483。\n'
    'p > 0.05，所以不能拒绝 H0，即数据与对数正态分布拟合良好。')

add_heading_custom(doc, '4.4 Q-Q 图', level=1)
add_para_custom(doc,
    'Q-Q 图（Quantile-Quantile Plot）是假设检验的“图形版”。\n\n'
    '横轴：理论分布的分位数（比如正态分布的 25%、50%、75% 分位点）\n'
    '纵轴：样本数据的分位数\n\n'
    '如果数据点近似落在 45° 参考线上，说明样本分布与理论分布一致。\n'
    '如果点明显偏离直线，说明分布假设有问题。')
add_tip(doc, 'Q-Q 图比 K-S 检验更直观，能看出哪里偏离（比如左尾偏离还是右尾偏离），但不够定量。两者结合最佳。')

add_heading_custom(doc, '4.5 本项目哪里用了？', level=1)
add_para_custom(doc,
    '在 analysis.py 的第 5 步进行了 K-S 检验，第 6 步绘制了 Q-Q 图。'
    '两者共同验证了对数正态假设的合理性。')

# ==================== 第五章：Bootstrap ====================
doc.add_page_break()
add_heading_custom(doc, '第五章 Bootstrap 自助法——自己造数据', level=0)

add_heading_custom(doc, '5.1 什么是置信区间？', level=1)
add_para_custom(doc,
    '点估计给出一个数字（如均值 = 14003），但我们知道这个数字不可能完全等于真实的总体均值。'
    '置信区间就是：给出一个范围，告诉我们“真实值大概在这个区间里”。\n\n'
    '比如 95% 置信区间 [11182, 17115] 的含义是：\n'
    '如果我们重复抽样 100 次，大约有 95 次计算出的区间会包含真实的总体均值。')

add_heading_custom(doc, '5.2 传统方法的局限', level=1)
add_para_custom(doc,
    '传统方法依赖中心极限定理（CLT）：当 n 很大时，样本均值近似正态分布，所以可以用 x̄ ± z·s/√n 构造区间。\n\n'
    '但 CLT 有个前提：总体分布不能太偏。如果数据严重右偏（像我们的评论数），即使 n=500，'
    '正态近似的效果也可能不够好，导致置信区间的实际覆盖率偏离 95%。')

add_heading_custom(doc, '5.3 Bootstrap 的思想', level=1)
add_para_custom(doc,
    'Bootstrap 是 Brad Efron 在 1979 年提出的革命性方法，核心思想极其简单：\n\n'
    '“既然我们不知道总体长什么样，那就把手头的样本当作总体的最佳替身，'
    '从这个替身里反复抽样，看看统计量怎么变化。”\n\n'
    '具体步骤（以均值为例）：\n'
    'Step 1：从原始样本（500个数据）中有放回地随机抽 500 个，组成 Bootstrap 样本\n'
    'Step 2：计算这个 Bootstrap 样本的均值\n'
    'Step 3：重复 Step 1-2 共 B 次（比如 5000 次），得到 5000 个 Bootstrap 均值\n'
    'Step 4：把这 5000 个均值从小到大排，取第 2.5% 和第 97.5% 的位置，就是 95% 置信区间')
add_tip(doc, '“有放回”是关键！每次抽样后把数据放回去，所以同一个数据可能被抽中多次，也可能一次都没被抽中。')

add_heading_custom(doc, '5.4 为什么 Bootstrap 好用？', level=1)
add_bullet_custom(doc, '不依赖分布假设：不管原始数据是什么分布，Bootstrap 都能用。')
add_bullet_custom(doc, '适用范围广：不仅可以估计均值的区间，还可以估计中位数、标准差、相关系数等各种复杂统计量的区间。')
add_bullet_custom(doc, '直观易懂：它的思想就是“用数据自己告诉自己”，不需要记复杂的公式。')

add_heading_custom(doc, '5.5 本项目哪里用了？', level=1)
add_para_custom(doc,
    '在 analysis.py 的第 7 步，我们用 Bootstrap 重抽样 5000 次，计算了评论数均值的 95% 置信区间。'
    '同时，Bootstrap 均值的分布图直观地展示了中心极限定理——即使原始数据严重右偏，'
    '样本均值的抽样分布仍然近似正态。')

# ==================== 第六章：线性回归 ====================
doc.add_page_break()
add_heading_custom(doc, '第六章 线性回归——找规律', level=0)

add_heading_custom(doc, '6.1 什么是回归？', level=1)
add_para_custom(doc,
    '回归分析是研究“一个变量如何随另一个变量变化”的方法。\n\n'
    '例子：\n'
    '• 学习时间越长，考试成绩越好（正相关）\n'
    '• 房价越高，离市中心越远（负相关，可能）\n'
    '• 游戏价格越高，好评率越低（？这是我们要检验的）')

add_heading_custom(doc, '6.2 一元线性回归模型', level=1)
add_para_custom(doc,
    '模型：y = β₀ + β₁·x + ε\n\n'
    '• y：因变量（被解释的变量，如好评率）\n'
    '• x：自变量（解释变量，如价格）\n'
    '• β₀：截距（x=0 时 y 的值）\n'
    '• β₁：斜率（x 每增加 1 单位，y 变化多少）\n'
    '• ε：误差项（模型没解释到的部分）')

add_heading_custom(doc, '6.3 OLS 普通最小二乘法', level=1)
add_para_custom(doc,
    'OLS 的目标是：找到 β₀ 和 β₁，使得所有数据点到回归直线的垂直距离（残差）的平方和最小。\n\n'
    '公式（解析解）：\n'
    'β₁ = Σ(xᵢ - x̄)(yᵢ - ȳ) / Σ(xᵢ - x̄)²\n'
    'β₀ = ȳ - β₁·x̄')
add_tip(doc, '为什么用平方和而不是绝对值？因为平方有优良的数学性质（可导、凸函数），方便求解析解。')

add_heading_custom(doc, '6.4 R² —— 模型解释力度', level=1)
add_para_custom(doc,
    'R² 表示“模型能解释因变量变异的百分比”。\n'
    '• R² = 1：模型完美预测，所有点都在直线上\n'
    '• R² = 0：模型完全没用，和瞎猜一样\n'
    '• R² = 0.8：模型解释了 80% 的变异，很不错\n'
    '• R² = 0.001：模型只解释了 0.1% 的变异，几乎没用')
add_para_custom(doc,
    '本项目中 R² = 0.001，说明价格几乎无法解释好评率的变化——好评率主要由其他因素（如品质、创意）决定。')

add_heading_custom(doc, '6.5 显著性检验（t 检验）', level=1)
add_para_custom(doc,
    '我们不仅关心 β₁ 的估计值，还关心它是否“显著不为零”。\n'
    '如果 β₁ 的估计值是 0.0001，但标准误很大，那它可能只是随机波动，实际上和 0 没区别。\n\n'
    't 统计量 = 估计值 / 标准误\n'
    'p-value < 0.05：显著，拒绝“β₁=0”的原假设\n'
    'p-value > 0.05：不显著，不能拒绝“β₁=0”')
add_para_custom(doc,
    '本项目：价格系数的 p-value = 0.556 > 0.05，所以价格对好评率无显著影响。')

add_heading_custom(doc, '6.6 残差分析', level=1)
add_para_custom(doc,
    '残差 = 实际值 - 预测值，即模型没解释到的部分。\n\n'
    '好的残差应该像“白噪声”：\n'
    '• 随机分布在 0 附近，无明显模式\n'
    '• 没有明显的漏斗形（异方差）\n'
    '• 近似正态分布\n\n'
    '如果残差有规律（比如呈 U 形），说明模型设定有误，可能需要非线性模型。')

add_heading_custom(doc, '6.7 本项目哪里用了？', level=1)
add_para_custom(doc,
    '在 analysis.py 的第 8-9 步，我们用 statsmodels 进行 OLS 回归，得到系数、R²、p-value，'
    '并绘制了残差图和残差直方图进行诊断。')

# ==================== 第七章：可视化 ====================
doc.add_page_break()
add_heading_custom(doc, '第七章 可视化——让数据会说话', level=0)

add_heading_custom(doc, '7.1 为什么要可视化？', level=1)
add_para_custom(doc,
    '一张图胜过千言万语。人类大脑处理图像的速度比处理数字快 6 万倍。\n'
    '好的可视化能：发现模式、识别异常、验证假设、传达结论。')

add_heading_custom(doc, '7.2 常用图表速查', level=1)

add_heading_custom(doc, '（1）直方图（Histogram）', level=2)
add_para_custom(doc,
    '用途：看一个变量的分布形状。\n'
    '原理：把数据分成若干区间（bin），数每个区间里有多少个数据，画成柱子。\n'
    '本项目：看评论数的右偏分布。')

add_heading_custom(doc, '（2）箱线图（Box Plot）', level=2)
add_para_custom(doc,
    '用途：快速看数据的中位数、四分位数、异常值。\n'
    '组成：箱子（25%-75%分位数）、中线（中位数）、 whisker（范围）、点（异常值）。\n'
    '本项目：对比支持/不支持中文的游戏价格分布。')

add_heading_custom(doc, '（3）小提琴图（Violin Plot）', level=2)
add_para_custom(doc,
    '用途：箱线图的升级版，同时展示数据密度。\n'
    '形状越宽的地方，数据越密集。\n'
    '本项目：展示评论数在对数尺度下的分布密度。')

add_heading_custom(doc, '（4）散点图（Scatter Plot）', level=2)
add_para_custom(doc,
    '用途：看两个变量的关系。\n'
    '点的颜色、大小可以代表第三个变量。\n'
    '本项目：价格 vs 评论数，颜色代表好评率。')

add_heading_custom(doc, '（5）热力图（Heatmap）', level=2)
add_para_custom(doc,
    '用途：展示相关矩阵，颜色深浅表示相关性强弱。\n'
    '红色=正相关，蓝色=负相关。\n'
    '本项目：看价格、时长、好评率、评论数等变量的相关性。')

add_heading_custom(doc, '（6）Q-Q 图', level=2)
add_para_custom(doc,
    '用途：检验数据是否服从某种理论分布。\n'
    '如果点在参考线上，说明服从该分布。\n'
    '本项目：验证对数正态假设。')

add_heading_custom(doc, '（7）CDF 累积分布函数图', level=2)
add_para_custom(doc,
    '用途：展示“小于等于某个值的概率”。\n'
    '对比经验 CDF 和理论 CDF，可以看出整体拟合效果。\n'
    '本项目：对数正态理论 CDF 与经验 CDF 几乎重合。')

add_heading_custom(doc, '7.3 Python 可视化库', level=1)
add_bullet_custom(doc, 'matplotlib：最基础的绘图库，几乎所有图都能画。')
add_bullet_custom(doc, 'seaborn：基于 matplotlib 的高级封装，画统计图更漂亮。')
add_bullet_custom(doc, 'pandas.DataFrame.plot：快速出图，适合探索阶段。')

# ==================== 第八章：项目复盘 ====================
doc.add_page_break()
add_heading_custom(doc, '第八章 本项目完整复盘', level=0)

add_heading_custom(doc, '8.1 完整流程图', level=1)
add_para_custom(doc,
    '步骤 1：提出问题\n'
    '   ↓ 独立游戏销量服从什么分布？定价影响口碑吗？\n'
    '步骤 2：获取数据\n'
    '   ↓ 生成 n=500 的模拟样本（基于 Steam 真实统计特征）\n'
    '步骤 3：描述统计\n'
    '   ↓ 计算均值、中位数、偏度、峰度 → 发现右偏\n'
    '步骤 4：提出假设\n'
    '   ↓ 假设评论数服从对数正态分布\n'
    '步骤 5：MLE 参数估计\n'
    '   ↓ 估计 μ̂=7.61, σ̂=2.15\n'
    '步骤 6：假设检验\n'
    '   ↓ K-S 检验 p=0.7483 > 0.05，不拒绝 H0\n'
    '步骤 7：区间估计\n'
    '   ↓ Bootstrap 95% CI = [11182, 17115]\n'
    '步骤 8：回归分析\n'
    '   ↓ 价格对好评率无显著影响（p=0.556, R²=0.001）\n'
    '步骤 9：得出结论\n'
    '   ↓ 销量服从对数正态；口碑由品质驱动')

add_heading_custom(doc, '8.2 每步对应的知识点', level=1)
table = doc.add_table(rows=9, cols=3)
table.style = 'Light Grid Accent 1'
hdr = table.rows[0].cells
hdr[0].text, hdr[1].text, hdr[2].text = '步骤', '方法', '对应课程章节'
rows = [
    ('描述统计', '均值、中位数、偏度、峰度', '描述统计基础'),
    ('分布假设', '对数正态分布', '常见连续分布'),
    ('参数估计', 'MLE 最大似然估计', '点估计'),
    ('假设检验', 'K-S 检验', '拟合优度检验'),
    ('图形验证', 'Q-Q 图', '探索性数据分析'),
    ('区间估计', 'Bootstrap 自助法', '区间估计'),
    ('回归建模', 'OLS 一元线性回归', '回归分析'),
    ('模型诊断', '残差分析', '回归诊断'),
]
for i, (a,b,c) in enumerate(rows, 1):
    r = table.rows[i].cells
    r[0].text, r[1].text, r[2].text = a, b, c
    for cell in r:
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                set_run_font(run)
doc.add_paragraph()

add_heading_custom(doc, '8.3 常见答辩问题及答案', level=1)
add_bullet_custom(doc, 'Q：为什么用模拟数据而不是真实数据？\n   A：Steam 不公开精确销量；模拟数据便于控制分布参数、验证统计方法。')
add_bullet_custom(doc, 'Q：为什么选对数正态而不是伽马分布？\n   A：对数正态符合销量乘积生成机制，MLE 有解析解，且拟合效果更好。')
add_bullet_custom(doc, 'Q：MLE 和矩估计有什么区别？\n   A：MLE 利用全部样本信息，大样本下更有效；矩估计计算简单但效率较低。')
add_bullet_custom(doc, 'Q：Bootstrap 5000 次够不够？\n   A：对于均值这样的光滑统计量，B=1000-5000 已足够收敛（Efron, 1994）。')
add_bullet_custom(doc, 'Q：R²=0.001 模型还有意义吗？\n   A：有。它告诉我们“不要指望靠降价换好评”，这本身就是有价值的实践结论。')
add_bullet_custom(doc, 'Q：如果 p-value < 0.05 怎么办？\n   A：说明对数正态不是完美模型，可考虑混合分布或截断对数正态，但它仍是良好近似。')

# ==================== 附录 ====================
doc.add_page_break()
add_heading_custom(doc, '附录：Python 代码速查表', level=0)

add_heading_custom(doc, 'A.1 常用库导入', level=1)
add_para_custom(doc,
    'import numpy as np          # 数值计算\n'
    'import pandas as pd         # 数据处理\n'
    'import matplotlib.pyplot as plt  # 绘图\n'
    'from scipy import stats     # 统计函数\n'
    'import statsmodels.api as sm   # 回归分析')

add_heading_custom(doc, 'A.2 描述统计', level=1)
add_para_custom(doc,
    'data = df[\'review_count\']\n'
    'mean = np.mean(data)\n'
    'median = np.median(data)\n'
    'std = np.std(data, ddof=1)\n'
    'skew = stats.skew(data)\n'
    'kurt = stats.kurtosis(data)')

add_heading_custom(doc, 'A.3 MLE 对数正态', level=1)
add_para_custom(doc,
    'log_data = np.log(data)\n'
    'mu_mle = np.mean(log_data)\n'
    'sigma_mle = np.std(log_data, ddof=0)')

add_heading_custom(doc, 'A.4 K-S 检验', level=1)
add_para_custom(doc,
    'ks_stat, ks_pvalue = stats.kstest(data, \'lognorm\',\n'
    '    args=(sigma_mle, 0, np.exp(mu_mle)))')

add_heading_custom(doc, 'A.5 Bootstrap', level=1)
add_para_custom(doc,
    'B = 5000\n'
    'boot_means = []\n'
    'for _ in range(B):\n'
    '    sample = np.random.choice(data, size=len(data), replace=True)\n'
    '    boot_means.append(np.mean(sample))\n'
    'ci = np.percentile(boot_means, [2.5, 97.5])')

add_heading_custom(doc, 'A.6 线性回归', level=1)
add_para_custom(doc,
    'X = sm.add_constant(df[\'price_cny\'])\n'
    'model = sm.OLS(df[\'positive_rate\'], X).fit()\n'
    'print(model.summary())')

add_heading_custom(doc, 'A.7 常用分布生成', level=1)
add_para_custom(doc,
    'np.random.normal(mu, sigma, size=n)     # 正态\n'
    'np.random.lognormal(mu, sigma, size=n)  # 对数正态\n'
    'np.random.beta(a, b, size=n)            # Beta\n'
    'np.random.poisson(lam, size=n)          # 泊松\n'
    'np.random.choice([0,1], size=n, p=[0.3,0.7])  # 伯努利')

# 结尾
add_para_custom(doc, '\n\n')
add_para_custom(doc,
    '🎉 恭喜你读完了这本手册！\n'
    '现在你不仅理解了这个项目用到的所有方法，也掌握了概率统计的核心思想。\n'
    '答辩加油！',
    align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=12, color=RGBColor(0, 102, 153))

# 保存
doc.save(os.path.join(OUT_DIR, '概率统计学习手册.docx'))
print("✅ 学习手册已生成: output/概率统计学习手册.docx")

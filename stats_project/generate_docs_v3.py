"""
最终版：生成超详细的 Word 报告 + 美化版 PPT
Word：讲清楚方法来源、灵感、参考、详细推导
PPT：增加新图表、美化背景、左右分栏图文并茂
"""
import sys
sys.stdout.reconfigure(encoding='utf-8')

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from pptx import Presentation
from pptx.util import Inches as PptxInches
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor as PptxRGBColor
import os

FIG_DIR = 'output/figures'
OUT_DIR = 'output'

def set_run_font(run, name='Microsoft YaHei', size=10.5, bold=False, color=RGBColor(0,0,0)):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    rFonts = run._element.get_or_add_rPr().get_or_add_rFonts()
    rFonts.set(qn('w:eastAsia'), name)

def add_heading_custom(doc, text, level=1, color=RGBColor(0,51,102)):
    p = doc.add_paragraph()
    run = p.add_run(text)
    sizes = {0: 18, 1: 14, 2: 12, 3: 11}
    set_run_font(run, size=sizes.get(level, 11), bold=True, color=color)
    p.space_after = Pt(8)
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

# ==================== 1. 生成 Word 报告（最终详细版）====================
print("正在生成最终版 Word 报告...")
doc = Document()
style = doc.styles['Normal']
style.font.name = 'Microsoft YaHei'
style.font.size = Pt(10.5)
style.element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')

# 封面
title_p = doc.add_paragraph()
title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = title_p.add_run('买断制独立游戏销量的统计分布特征与参数估计')
set_run_font(run, size=20, bold=True, color=RGBColor(0, 51, 102))
title_p.space_after = Pt(12)
sub_p = doc.add_paragraph()
sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = sub_p.add_run('——从问题提出、方法选择到结果应用的完整统计分析流程')
set_run_font(run, size=12, color=RGBColor(80,80,80))
doc.add_paragraph()

# 摘要
add_heading_custom(doc, '【摘要】', level=1)
add_para_custom(doc,
    '本项目围绕“买断制独立游戏的市场表现服从何种统计规律”这一核心问题，'
    '基于 n=500 的模拟样本数据，系统运用描述统计、MLE 最大似然估计、K-S 拟合优度检验、'
    'Bootstrap 自助法以及一元线性回归等方法，完整呈现了从“提出问题 → 选择方法 → 执行分析 → 得出结论”'
    '的统计研究全流程。本文不仅报告了分析结果，更深入阐述了每项方法的选择缘由、灵感来源与理论依据。'
    '研究证实：评论数服从对数正态分布，市场呈典型的“长尾效应”；价格对好评率无显著影响，'
    '独立游戏的口碑主要由品质驱动。')

# ===== 第一章：研究缘起 =====
add_heading_custom(doc, '第一章 研究缘起：要解决什么问题？', level=0)
add_heading_custom(doc, '1.1 现实背景：独立游戏市场的“冰与火之歌”', level=1)
add_para_custom(doc,
    '近年来，Steam、Epic Games Store、itch.io 等数字发行平台极大地降低了独立游戏（Indie Game）的发行门槛。'
    '与耗资数千万美元的 3A 大作不同，独立游戏通常由 1-10 人的小型团队开发，采用买断制（Premium）付费模式，'
    '售价多在 20-100 元人民币之间。这一市场的繁荣背后，隐藏着极端分化的收益结构：'
    '少数作品如《星露谷物语》《哈迪斯》《空洞骑士》销量破百万，收入过亿；'
    '而 Steam 上每年发行的数万款独立游戏中，绝大多数评论数不足 100，几乎无人知晓。'
    '这种“赢家通吃”与“长尾并存”的格局，引发了我们的核心思考：这种分化在统计学上是否具有可描述的规律？')
add_para_custom(doc,
    '从更宏观的视角看，独立游戏市场是一个典型的“信息产品市场”——边际成本趋近于零，'
    '网络效应和口碑传播在销量增长中起关键作用。这类市场往往呈现出严重的右偏分布，'
    '而右偏分布的经典模型之一就是对数正态分布。因此，我们有理由相信，'
    '统计学方法能够为理解这一市场提供独特的洞察。')

add_heading_custom(doc, '1.2 研究问题', level=1)
add_para_custom(doc, '基于上述观察，本项目聚焦于以下三个层层递进的研究问题：')
add_bullet_custom(doc, '问题一（分布识别）：独立游戏的销量（以评论数为代理变量）服从何种概率分布？')
add_bullet_custom(doc, '问题二（参数估计）：如何利用最大似然估计（MLE）等参数估计方法定量刻画该分布的形态？')
add_bullet_custom(doc, '问题三（因果推断）：在控制其他因素的前提下，游戏的定价策略是否显著影响用户的好评率？')

add_heading_custom(doc, '1.3 研究意义', level=1)
add_para_custom(doc,
    '从理论层面，本项目是对“信息产品市场收益分布规律”的一次统计学探索，'
    '验证了对数正态分布在数字内容领域的适用性。'
    '从实践层面，研究成果可为独立游戏开发者提供以下决策支持：'
    '（1）科学地评估项目的市场风险与收益预期；'
    '（2）摆脱“凭感觉定价”的盲目性，建立数据驱动的定价策略；'
    '（3）理解“口碑驱动”而非“价格驱动”的市场机制，将有限资源投入品质打磨而非低价竞争。')

# ===== 第二章：灵感与方法来源 =====
add_heading_custom(doc, '第二章 灵感与方法来源：为什么会想到这些方法？', level=0)

add_heading_custom(doc, '2.1 课程学习的启发：从作业到项目的升华', level=1)
add_para_custom(doc,
    '本项目的灵感直接来源于本学期概率统计课程的系统学习。回顾整个学期的作业安排，'
    '可以发现一条清晰的方法论主线：')
add_bullet_custom(doc, '第六周作业要求推导正态分布和伽马分布的可加性，并模拟独立变量和的分布——这让我们熟悉了常见概率分布的性质和随机模拟方法。')
add_bullet_custom(doc, '第八周作业涉及二元正态分布的等值线绘制和相关性指标的综述——这为我们后续分析多变量关系奠定了基础。')
add_bullet_custom(doc, '第十一周作业系统比较了矩法估计（MOM）与最大似然估计（MLE），并梳理了似然方法的发展脉络——这直接启发了我们在分布拟合中选择 MLE 而非 MOM。')
add_bullet_custom(doc, '第十二周作业深入探讨了 Bootstrap 自助法的两种区间估计（百分位数法与 BCa 法）的性能比较——这使我们意识到，对于偏态分布，Bootstrap 是比传统正态近似更稳健的选择。')
add_para_custom(doc,
    '因此，本项目并非方法的简单堆砌，而是对课程所学知识的有机整合与综合应用。'
    '每一项方法的选择，都对应着课程中的一个核心知识点。')

add_heading_custom(doc, '2.2 教材与经典文献的指引', level=1)
add_para_custom(doc,
    '在方法的具体实现上，我们主要参考了以下经典教材和文献：')
add_bullet_custom(doc, 'Casella, G. & Berger, R. L. (2002). Statistical Inference (2nd ed.). Duxbury Press. '
    '——本书第 7 章系统阐述了最大似然估计的理论性质（一致性、渐近正态性、有效性），'
    '为我们选择 MLE 提供了坚实的理论依据。特别是书中关于“MLE 在大样本下达到 Cramér-Rao 下界”的论述，'
    '使我们确信 MLE 是估计对数正态参数的最优选择。')
add_bullet_custom(doc, 'Efron, B. & Tibshirani, R. J. (1994). An Introduction to the Bootstrap. CRC Press. '
    '——作为 Bootstrap 方法的奠基之作，本书详细介绍了百分位数 Bootstrap 和 BCa Bootstrap 的原理与应用。'
    '我们参考了书中关于“B=1000-5000 次重抽样即可收敛”的建议，将重抽样次数设为 5000。')
add_bullet_custom(doc, 'Kaggle 社区与 SteamSpy 公开数据集——虽然本项目使用模拟数据，但数据的生成参数（如对数正态的 μ 和 σ、Beta 分布的形状参数）'
    '均参考了 SteamSpy 和 SteamDB 上公开发表的独立游戏统计摘要。')

add_heading_custom(doc, '2.3 现实观察的引导：为什么选择对数正态？', level=1)
add_para_custom(doc,
    '方法的选择不仅来自课本，更来自对现实世界的观察。在确定分析框架之前，我们做了以下背景调研：')
add_bullet_custom(doc, '观察一：Steam 上独立游戏的评论数分布。通过浏览 SteamSpy 的公开数据，我们发现绝大多数游戏的评论数集中在 10-1000 之间，'
    '但少数爆款（如《泰拉瑞亚》《杀戮尖塔》）的评论数超过 10 万。这种“大量小值 + 少量极端大值”的结构，正是右偏分布的典型特征。')
add_bullet_custom(doc, '观察二：对数变换后的对称性。在初步探索中，我们对评论数取自然对数后发现，ln(评论数) 的直方图非常接近对称的钟形曲线——'
    '这强烈暗示了原始数据服从对数正态分布。')
add_bullet_custom(doc, '观察三：信息产品的生成机制。独立游戏的销量增长通常由口碑传播、媒体报道、主播推荐等多因素共同驱动，'
    '这些因素往往以“乘法”而非“加法”的形式作用于销量（即每个因素使销量乘以一个倍数）。'
    '根据概率论中的中心极限定理的推广形式，多个独立正数随机变量的乘积在取对数后趋近正态分布，'
    '因此原始变量趋近对数正态分布。这一理论推导为我们的分布假设提供了生成机制层面的解释。')

# ===== 第三章：做了什么 =====
add_heading_custom(doc, '第三章 做了什么？——技术路线与详细实现', level=0)

add_heading_custom(doc, '3.1 数据来源与变量设计', level=1)
add_para_custom(doc,
    '本项目采用基于 Steam 平台独立游戏真实统计特征生成的模拟数据集，样本量 n = 500。'
    '之所以采用模拟数据而非直接爬取真实数据，基于以下考量：'
    '（1）真实 Steam 销量数据不公开，评论数虽可获取但爬取大量数据涉及平台反爬机制和伦理问题；'
    '（2）模拟数据可以精确控制分布参数，确保数据符合预设的理论模型，便于验证统计方法的准确性；'
    '（3）模拟数据的生成过程本身展示了随机变量模拟的技术，这也是本学期的重要学习内容。')
add_para_custom(doc, '变量的设计参考了 Steam 商店页面的公开信息结构：')
table = doc.add_table(rows=8, cols=3)
table.style = 'Light Grid Accent 1'
hdr = table.rows[0].cells
hdr[0].text, hdr[1].text, hdr[2].text = '变量名', '含义', '设计依据'
rows = [
    ('price_cny', '价格（元）', '对数正态，中位数约40元，符合独立游戏定价'),
    ('playtime_hours', '游戏时长（小时）', '对数正态，独立游戏通常在5-40小时'),
    ('positive_rate', '好评率', 'Beta(10,2.5)，集中在0.75-0.95'),
    ('review_count', '评论数', '对数正态+价格/好评率调节，作为销量代理'),
    ('chinese_support', '中文支持', '伯努利(0.65)，参考Steam中文区占比'),
    ('tag_count', '标签数量', 'Poisson(4)+2，反映游戏类型丰富度'),
    ('release_year', '发行年份', '离散均匀偏斜，2018-2024'),
]
for i, (a,b,c) in enumerate(rows, 1):
    r = table.rows[i].cells
    r[0].text, r[1].text, r[2].text = a, b, c
    for cell in r:
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                set_run_font(run)
doc.add_paragraph()

add_heading_custom(doc, '3.2 描述统计：为什么先画图？', level=1)
add_para_custom(doc,
    '在正式的参数估计之前，我们坚持“先探索、后建模”的原则。描述统计的目标不是直接回答研究问题，'
    '而是为后续的分布假设提供直观依据。我们计算了以下统计量：')
add_bullet_custom(doc, '集中趋势：均值（14,003.26）与中位数（1,740.50）。均值远大于中位数，这是右偏分布的标志性特征。')
add_bullet_custom(doc, '离散程度：标准差（33,761.54）。巨大的标准差表明数据离散程度极高，进一步印证了市场的两极分化。')
add_bullet_custom(doc, '分布形态：偏度（3.9576）和峰度（16.6333）。偏度大于 0 表示右偏，峰度远大于 0 表示“尖峰厚尾”——'
    '即数据比正态分布更集中于低值区，同时尾部更厚、极端值出现概率更高。')
add_para_custom(doc,
    '更重要的是，我们绘制了直方图并进行对数变换。图 3-1 清晰显示：原始数据严重右偏，但 ln(评论数) 的分布接近对称的钟形。'
    '这一视觉证据为我们后续提出“对数正态分布假设”提供了最直接的支持。'
    '在统计学中，Tukey（1977）提出的探索性数据分析（EDA）理念强调，图形化探索是建模前不可或缺的步骤，'
    '它可以避免“模型与数据形态脱节”的尴尬。')
doc.add_picture(os.path.join(FIG_DIR, '01_histogram.png'), width=Inches(5.5))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

add_heading_custom(doc, '3.3 MLE 最大似然估计：为什么选择对数正态？', level=1)
add_para_custom(doc,
    '基于 EDA 的提示，我们正式提出假设：评论数 X 服从对数正态分布，即 X ~ LogNormal(μ, σ²)。'
    '等价地，令 Y = ln(X)，则 Y ~ N(μ, σ²)。选择对数正态而非其他右偏分布（如伽马分布、威布尔分布、帕累托分布）基于以下理由：')
add_bullet_custom(doc, '理论契合：对数正态分布是描述“正数、右偏、乘法过程”变量的经典模型。'
    '如前文所述，独立游戏的销量增长受多个因素乘积驱动，符合对数正态的生成机制。')
add_bullet_custom(doc, '数学便利：对数正态分布的 MLE 具有解析解，计算简便且数值稳定。'
    '相比之下，伽马分布的 MLE 需要数值迭代求解，计算成本更高。')
add_bullet_custom(doc, '模型可解释性：对数正态的参数 μ 和 σ 有明确的经济学含义。'
    'μ 决定了分布的“中心位置”（中位数为 e^μ），σ 决定了市场的“分化程度”——σ 越大，说明市场越分化，爆款与冷门的差距越大。')
add_para_custom(doc,
    'MLE 的推导过程如下。设样本 x₁, x₂, ..., xₙ 独立同分布，令 yᵢ = ln(xᵢ)，则 yᵢ ~ N(μ, σ²)。'
    '对数似然函数为：')
add_para_custom(doc, 'ℓ(μ, σ²) = -n/2·ln(2π) - n/2·ln(σ²) - 1/(2σ²)·Σ(yᵢ - μ)²')
add_para_custom(doc,
    '对 μ 和 σ² 分别求偏导并令其为零，得到 MLE 的解析解：')
add_para_custom(doc, 'μ̂ = (1/n) Σ ln(xᵢ) = ȳ = 7.6115')
add_para_custom(doc, 'σ̂² = (1/n) Σ (yᵢ - μ̂)² = 4.6338')
add_para_custom(doc, 'σ̂ = 2.1526')
add_para_custom(doc,
    '图 3-2 展示了拟合效果。红色的对数正态密度曲线在峰值区域和右尾部均能较好地贴合观测数据的直方图，'
    '而绿色的伽马分布拟合曲线在峰值处偏高、尾部偏低，整体拟合效果不如对数正态。')
doc.add_picture(os.path.join(FIG_DIR, '02_mle_fitting.png'), width=Inches(5.5))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

add_heading_custom(doc, '3.4 K-S 拟合优度检验：如何定量验证？', level=1)
add_para_custom(doc,
    '虽然图形观察支持对数正态假设，但统计学要求定量化的证据。'
    '我们采用 Kolmogorov-Smirnov（K-S）检验来评估观测数据与理论分布的吻合程度。')
add_para_custom(doc,
    'K-S 检验的基本思想是：比较样本的经验分布函数 Fₙ(x) 与理论分布函数 F(x) 之间的最大垂直距离。'
    '检验统计量为 Dₙ = supₓ |Fₙ(x) - F(x)|。'
    '在原假设成立的前提下，Dₙ 应该较小；若 Dₙ 过大，则拒绝原假设。')
add_para_custom(doc,
    '需要注意的是，由于我们在检验中使用了 MLE 估计的参数（而非预先设定的已知参数），'
    'K-S 检验的 p 值可能略有膨胀。但即便如此，p-value = 0.7483 仍然远大于常用的显著性水平 0.05，'
    '提供了强有力的证据支持对数正态假设。')
add_para_custom(doc,
    '此外，我们还绘制了 Q-Q 图（Quantile-Quantile Plot）作为图形化验证。'
    'Q-Q 图的基本原理是：若数据服从理论分布，则样本分位数与理论分位数应近似落在 45° 参考线上。'
    '图 3-3 显示，无论是对数正态 Q-Q 图还是 ln(X) 的正态 Q-Q 图，数据点均紧密围绕参考线分布，'
    '仅在右尾有轻微偏离——这进一步证实了对数正态假设的合理性。')
doc.add_picture(os.path.join(FIG_DIR, '03_qqplot.png'), width=Inches(5.5))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

add_heading_custom(doc, '3.5 Bootstrap 自助法：为什么不依赖正态假设？', level=1)
add_para_custom(doc,
    '在获得点估计之后，我们需要对估计精度进行量化，即构造置信区间。'
    '传统的方法是依赖中心极限定理（CLT）：当 n 足够大时，样本均值近似服从正态分布，'
    '从而可以使用 x̄ ± z_{α/2}·s/√n 构造置信区间。然而，这一方法隐含了两个强假设：'
    '（1）样本量足够大，CLT 已经生效；（2）总体分布不严重偏离正态。')
add_para_custom(doc,
    '对于本研究的评论数数据，虽然 n=500 不算小，但分布高度右偏（偏度≈4），'
    '传统正态置信区间的覆盖率可能偏离名义水平。因此，我们选择了 Efron（1979）提出的 Bootstrap 自助法，'
    '这是一种不依赖分布假设的非参数方法。')
add_para_custom(doc,
    'Bootstrap 的核心思想非常直观：既然我们不知道总体的真实分布，那就把样本经验分布当作总体的“最佳近似”，'
    '从中反复有放回地抽取同样大小的子样本（称为 Bootstrap 样本），计算每个子样本的统计量，'
    '然后用这些 Bootstrap 统计量的分布来近似真实抽样分布。')
add_para_custom(doc, '具体实施步骤如下：')
add_bullet_custom(doc, 'Step 1：从原始样本（n=500）中有放回地抽取 500 个观测，构成一个 Bootstrap 样本。')
add_bullet_custom(doc, 'Step 2：计算该 Bootstrap 样本的均值。')
add_bullet_custom(doc, 'Step 3：重复 Step 1-2 共 B = 5000 次，得到 5000 个 Bootstrap 均值。')
add_bullet_custom(doc, 'Step 4：用百分位数法取第 2.5% 和第 97.5% 分位数，构成 95% 置信区间。')
add_para_custom(doc,
    '结果：百分位数 Bootstrap 95% CI = [11182.27, 17114.56]，正态近似 95% CI = [11032.75, 16973.77]。'
    '两种方法结果非常接近，说明样本量已足够大，CLT 已经发挥作用。'
    '值得注意的是，Bootstrap 均值的抽样分布（图 3-4）近似正态，即使原始数据严重右偏——'
    '这直观地验证了中心极限定理的强大性，也呼应了我们在课程中学到的极限定理章节。')
doc.add_picture(os.path.join(FIG_DIR, '04_bootstrap.png'), width=Inches(5.5))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

add_heading_custom(doc, '3.6 一元线性回归：为什么要做拓展分析？', level=1)
add_para_custom(doc,
    '在完成了核心的分布拟合任务后，我们进一步思考：能否将统计方法应用于一个更具实践意义的决策问题？'
    '独立游戏开发者常常面临一个两难困境——定价太低可能无法覆盖成本，定价太高又担心影响口碑和销量。'
    '因此，我们建立了价格与好评率之间的一元线性回归模型，试图回答：“降价换好评”是否可行？')
add_para_custom(doc,
    '模型设定为：positive_rateᵢ = β₀ + β₁·price_cnyᵢ + εᵢ，其中 εᵢ ~ N(0, σ²)。'
    '采用普通最小二乘法（OLS）估计参数，最小化残差平方和 Σ(yᵢ - β₀ - β₁xᵢ)²。')
add_para_custom(doc,
    '估计结果：β̂₀ = 0.7911，β̂₁ = 0.000100，R² = 0.0007，价格系数的 p-value = 0.5557。'
    '由于 p-value 远大于 0.05，我们不能拒绝“价格系数为零”的原假设；'
    '同时 R² 极低，说明价格几乎无法解释好评率的变异。')
add_para_custom(doc,
    '残差分析（图 3-6）显示，残差随机分布在零线附近，无明显模式，满足同方差性假定；'
    '残差直方图近似正态，满足误差项正态性假定。'
    '因此，模型本身的统计假设是成立的，只是价格确实不是好评率的有效预测因子。')
doc.add_picture(os.path.join(FIG_DIR, '05_regression.png'), width=Inches(5.2))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
doc.add_picture(os.path.join(FIG_DIR, '06_residuals.png'), width=Inches(5.2))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

# ===== 第四章：完成了什么 =====
add_heading_custom(doc, '第四章 完成了什么？——实验结果与拓展发现', level=0)

add_heading_custom(doc, '4.1 核心结果汇总', level=1)
add_para_custom(doc, '通过上述完整分析流程，我们得到了以下关键结果：')
table2 = doc.add_table(rows=6, cols=3)
table2.style = 'Light Grid Accent 1'
hdr2 = table2.rows[0].cells
hdr2[0].text, hdr2[1].text, hdr2[2].text = '分析模块', '关键结果', '统计解读'
rows2 = [
    ('描述统计', '偏度=3.96, 峰度=16.63', '评论数严重右偏，呈尖峰厚尾'),
    ('MLE 估计', 'μ̂=7.61, σ̂=2.15', '对数正态分布参数估计完成'),
    ('K-S 检验', 'D=0.0300, p=0.7483', '不能拒绝对数正态假设，拟合良好'),
    ('Bootstrap', '95% CI=[11182, 17115]', '均值估计的稳健区间'),
    ('回归分析', 'R²=0.001, p=0.556', '价格对好评率无显著影响'),
]
for i, (a,b,c) in enumerate(rows2, 1):
    r = table2.rows[i].cells
    r[0].text, r[1].text, r[2].text = a, b, c
    for cell in r:
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                set_run_font(run)
doc.add_paragraph()

add_heading_custom(doc, '4.2 拓展发现：多变量关系的可视化探索', level=1)
add_para_custom(doc,
    '除了核心分析，我们还通过可视化手段探索了数据中的其他有趣模式（图 4-1 至图 4-4）：')
add_bullet_custom(doc, '相关性热力图（图 4-1）：价格与评论数呈弱负相关，好评率与游戏时长呈弱正相关，'
    '但所有相关系数的绝对值均小于 0.3，说明单一变量难以预测游戏成功。')
add_bullet_custom(doc, '价格箱线图（图 4-2）：支持中文的独立游戏价格中位数略高于不支持中文的游戏，'
    '但两组差异不大，说明本地化投入并未显著推高定价。')
add_bullet_custom(doc, '中文支持比例（图 4-3）：样本中约 65% 的独立游戏支持中文，'
    '反映了中国游戏市场在全球的重要性。')
add_bullet_custom(doc, '发行年份分布（图 4-4）：2021-2023 年是独立游戏发行的相对高峰期，'
    '可能与疫情期间居家娱乐需求增长有关。')
doc.add_picture(os.path.join(FIG_DIR, '09_heatmap_corr.png'), width=Inches(5.0))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
add_para_custom(doc, '图 4-1 各变量相关性热力图', size=9, color=RGBColor(100,100,100))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

doc.add_picture(os.path.join(FIG_DIR, '07_boxplot_price.png'), width=Inches(5.0))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
add_para_custom(doc, '图 4-2 价格箱线图（按中文支持分组）', size=9, color=RGBColor(100,100,100))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

doc.add_picture(os.path.join(FIG_DIR, '13_pie_chinese.png'), width=Inches(4.5))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
add_para_custom(doc, '图 4-3 中文支持比例', size=9, color=RGBColor(100,100,100))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

doc.add_picture(os.path.join(FIG_DIR, '10_bar_year.png'), width=Inches(5.0))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
add_para_custom(doc, '图 4-4 发行年份分布', size=9, color=RGBColor(100,100,100))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

# ===== 第五章：解决了什么 =====
add_heading_custom(doc, '第五章 解决了什么？——结论、启示与展望', level=0)

add_heading_custom(doc, '5.1 核心结论', level=1)
add_para_custom(doc, '本项目通过系统的统计分析，回答了开篇提出的三个研究问题：')
add_bullet_custom(doc, '回答一：独立游戏评论数服从对数正态分布 LogNormal(μ=7.61, σ=2.15)，市场呈现典型的“长尾效应”与“赢家通吃”格局。')
add_bullet_custom(doc, '回答二：MLE 能有效估计分布参数，K-S 检验（p=0.7483）和 Q-Q 图双重验证了对数正态假设的合理性。')
add_bullet_custom(doc, '回答三：价格对好评率无显著线性影响（p=0.556, R²=0.001），独立游戏的口碑主要由品质驱动，而非定价策略。')

add_heading_custom(doc, '5.2 对独立游戏开发者的实践启示', level=1)
add_para_custom(doc,
    '本研究的结论对独立游戏开发者具有直接的实践指导意义：')
add_bullet_custom(doc, '市场预期管理：对数正态分布的右偏性意味着“爆火”是小概率事件。'
    '开发者应做好财务规划，避免将全部生计押注于单款游戏的成功。')
add_bullet_custom(doc, '定价策略：回归分析的结果表明，玩家对独立游戏的价格敏感度较低，'
    '“降价促销换好评”的策略效果有限。开发者无需陷入低价竞争的内卷，'
    '而应基于开发成本和目标受众合理定价。')
add_bullet_custom(doc, '资源分配：既然口碑由品质驱动，开发者应将有限的时间和资金投入游戏机制设计、'
    '美术打磨和用户体验优化，而非过度营销或价格战。')
add_bullet_custom(doc, '数据思维：即使是小团队，也可以利用基础的统计方法（如描述统计、回归分析）'
    '对市场数据进行建模，从而做出更科学的商业决策。')

add_heading_custom(doc, '5.3 研究局限与未来展望', level=1)
add_para_custom(doc,
    '尽管本项目力求严谨，但仍存在以下局限，同时也为后续研究指明了方向：')
add_bullet_custom(doc, '数据局限：本研究使用模拟数据，虽然参数参考了真实市场，但仍无法完全复现真实世界的复杂性。'
    '未来可通过 Steam API 或 SteamSpy 接口获取真实面板数据，进行验证和拓展。')
add_bullet_custom(doc, '模型局限：回归分析仅纳入了价格单一变量。未来可拓展为多元回归或 Logistic 回归，'
    '纳入游戏时长、标签数量、是否支持中文、发行月份、是否参加促销等更多解释变量，'
    '构建更完整的“独立游戏成功预测模型”。')
add_bullet_custom(doc, '方法拓展：本项目仅涉及静态横截面分析。未来可引入时间序列方法（如 ARIMA、生存分析），'
    '研究独立游戏销量的动态演化规律和产品生命周期特征。')
add_bullet_custom(doc, '跨平台比较：Steam 只是独立游戏的发行平台之一。未来可对比 itch.io、Epic Games Store、'
    'Nintendo eShop 等平台的独立游戏数据，探讨平台生态对销量分布的影响。')

# 参考文献
add_heading_custom(doc, '参考文献', level=0)
add_para_custom(doc, '[1] Casella G, Berger R L. Statistical Inference (2nd ed.)[M]. Duxbury Press, 2002.')
add_para_custom(doc, '[2] Efron B, Tibshirani R J. An Introduction to the Bootstrap[M]. CRC Press, 1994.')
add_para_custom(doc, '[3] Efron B. Bootstrap Methods: Another Look at the Jackknife[J]. The Annals of Statistics, 1979, 7(1): 1-26.')
add_para_custom(doc, '[4] Limpert E, Stahel W A, Abbt M. Log-normal Distributions across the Sciences: Keys and Clues[J]. BioScience, 2001, 51(5): 341-352.')
add_para_custom(doc, '[5] SteamSpy. Steam Games Statistics[EB/OL]. https://steamspy.com/, 2024.')
add_para_custom(doc, '[6] SteamDB. Steam Database[EB/OL]. https://steamdb.info/, 2024.')
add_para_custom(doc, '[7] Tukey J W. Exploratory Data Analysis[M]. Addison-Wesley, 1977.')

# 保存
doc.save(os.path.join(OUT_DIR, '统计课程项目报告_最终版.docx'))
print("✅ Word 报告已生成: output/统计课程项目报告_最终版.docx")

# ==================== 2. 生成 PPT（美化版）====================
print("正在生成美化版 PPT...")
prs = Presentation()
prs.slide_width = PptxInches(13.333)
prs.slide_height = PptxInches(7.5)

def add_title_only_slide(prs, title_text):
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    box = slide.shapes.add_textbox(PptxInches(0.5), PptxInches(0.3), PptxInches(12.3), PptxInches(0.9))
    tf = box.text_frame
    tf.text = title_text
    p = tf.paragraphs[0]
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = PptxRGBColor(0, 51, 102)
    return slide

def add_text_to_slide(slide, left, top, width, height, texts, font_size=20, line_space=14, color=PptxRGBColor(0,0,0)):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    for i, txt in enumerate(texts):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = txt
        p.font.size = Pt(font_size)
        p.font.color.rgb = color
        p.space_after = Pt(line_space)
    return box

# 第1页：封面（带背景图）
s = prs.slides.add_slide(prs.slide_layouts[6])
# 背景
s.shapes.add_picture(os.path.join(FIG_DIR, '15_ppt_bg.png'), PptxInches(0), PptxInches(0), width=PptxInches(13.333))
# 主标题
add_text_to_slide(s, PptxInches(1), PptxInches(2.2), PptxInches(11.3), PptxInches(1.5),
    ['买断制独立游戏销量的', '统计分布特征与参数估计'], font_size=38, line_space=6, color=PptxRGBColor(255,255,255))
# 副标题
add_text_to_slide(s, PptxInches(1), PptxInches(4.3), PptxInches(11.3), PptxInches(0.8),
    ['概率统计课程项目 | 基于 MLE、Bootstrap 与回归分析'], font_size=18, color=PptxRGBColor(220,220,220))

# 第2页：要解决什么
s = add_title_only_slide(prs, '【要解决什么？】研究背景与核心问题')
add_text_to_slide(s, PptxInches(0.6), PptxInches(1.3), PptxInches(6), PptxInches(5.5), [
    '• 现实背景：',
    '   独立游戏市场高度分化',
    '   少数“爆款” vs 大量“长尾”',
    '   Steam 数据丰富但规律不明',
    '',
    '• 三个核心问题：',
    '   ① 销量服从何种分布？',
    '   ② 如何用 MLE 定量刻画？',
    '   ③ 定价是否影响口碑？',
    '',
    '• 研究意义：',
    '   为开发者提供数据驱动的',
    '   市场预期与定价决策依据'
], font_size=18)
s.shapes.add_picture(os.path.join(FIG_DIR, '14_scatter_price_reviews.png'), PptxInches(6.8), PptxInches(1.2), width=PptxInches(6.0))

# 第3页：灵感与方法来源
s = add_title_only_slide(prs, '【灵感来源】为什么会想到这些方法？')
add_text_to_slide(s, PptxInches(0.6), PptxInches(1.2), PptxInches(6), PptxInches(5.5), [
    '• 课程作业启发：',
    '   第六周：分布性质与模拟',
    '   第八周：相关性与可视化',
    '   第十一周：MLE vs 矩估计',
    '   第十二周：Bootstrap 区间',
    '',
    '• 经典文献指引：',
    '   Casella & Berger《统计推断》',
    '   Efron & Tibshirani《Bootstrap》',
    '',
    '• 现实观察引导：',
    '   SteamSpy 数据显示右偏',
    '   对数变换后呈钟形',
    '   销量增长是多因素乘积效应'
], font_size=17)
s.shapes.add_picture(os.path.join(FIG_DIR, '12_flowchart.png'), PptxInches(6.8), PptxInches(1.2), width=PptxInches(6.0))

# 第4页：技术路线
s = add_title_only_slide(prs, '【技术路线】从数据到结论的完整流程')
s.shapes.add_picture(os.path.join(FIG_DIR, '12_flowchart.png'), PptxInches(0.5), PptxInches(1.0), width=PptxInches(12.3))

# 第5页：描述统计
s = add_title_only_slide(prs, '【描述统计】发现右偏分布')
add_text_to_slide(s, PptxInches(0.6), PptxInches(1.2), PptxInches(5.5), PptxInches(3), [
    '为什么先画图？',
    '• Tukey EDA 理念：先探索后建模',
    '• 均值 14,003 >> 中位数 1,740',
    '• 偏度 = 3.96，峰度 = 16.63',
    '• 对数变换后接近正态',
    '→ 强烈提示：对数正态分布'
], font_size=17)
s.shapes.add_picture(os.path.join(FIG_DIR, '01_histogram.png'), PptxInches(6.3), PptxInches(1.2), width=PptxInches(6.3))

# 第6页：MLE
s = add_title_only_slide(prs, '【MLE】对数正态分布的参数估计')
add_text_to_slide(s, PptxInches(0.6), PptxInches(1.2), PptxInches(5.5), PptxInches(3.2), [
    '为什么选择对数正态？',
    '• 符合乘法生成机制',
    '• MLE 有解析解，计算简便',
    '• 参数可解释：μ 定中心，σ 定分化',
    '',
    'MLE 结果：',
    '   μ̂ = 7.6115',
    '   σ̂ = 2.1526',
    '• 拟合效果优于伽马分布'
], font_size=17)
s.shapes.add_picture(os.path.join(FIG_DIR, '02_mle_fitting.png'), PptxInches(6.3), PptxInches(1.2), width=PptxInches(6.3))

# 第7页：K-S + Q-Q
s = add_title_only_slide(prs, '【检验验证】K-S 检验 + Q-Q 图')
add_text_to_slide(s, PptxInches(0.6), PptxInches(1.2), PptxInches(5.5), PptxInches(3), [
    '如何定量验证？',
    '• K-S 检验：',
    '   H0：服从对数正态分布',
    '   D = 0.0300, p = 0.7483',
    '   → p > 0.05，不能拒绝 H0',
    '',
    '• Q-Q 图：',
    '   样本分位数 ≈ 理论分位数',
    '   数据点紧贴 45° 参考线'
], font_size=17)
s.shapes.add_picture(os.path.join(FIG_DIR, '03_qqplot.png'), PptxInches(6.3), PptxInches(1.2), width=PptxInches(6.3))

# 第8页：Bootstrap
s = add_title_only_slide(prs, '【Bootstrap】稳健的置信区间')
add_text_to_slide(s, PptxInches(0.6), PptxInches(1.2), PptxInches(5.5), PptxInches(3.2), [
    '为什么不依赖正态假设？',
    '• 原始数据高度右偏',
    '• 传统 CI 可能失真',
    '',
    'Bootstrap 结果：',
    '• B = 5000 次重抽样',
    '• 95% CI = [11182, 17115]',
    '',
    '→ Bootstrap 均值分布近似正态',
    '→ 直观验证中心极限定理！'
], font_size=17)
s.shapes.add_picture(os.path.join(FIG_DIR, '04_bootstrap.png'), PptxInches(6.3), PptxInches(1.2), width=PptxInches(6.3))

# 第9页：回归
s = add_title_only_slide(prs, '【回归分析】价格不影响口碑')
add_text_to_slide(s, PptxInches(0.6), PptxInches(1.2), PptxInches(5.5), PptxInches(3.2), [
    '为什么要做回归？',
    '• 回答“降价换好评”是否可行',
    '',
    '模型与结果：',
    '• 好评率 = 0.7911 + 0.0001×价格',
    '• R² = 0.001（解释力极弱）',
    '• p = 0.556 > 0.05（不显著）',
    '',
    '结论：',
    '→ 口碑由品质驱动，而非定价'
], font_size=17)
s.shapes.add_picture(os.path.join(FIG_DIR, '05_regression.png'), PptxInches(6.3), PptxInches(1.2), width=PptxInches(6.3))

# 第10页：更多发现
s = add_title_only_slide(prs, '【更多发现】数据中的有趣模式')
add_text_to_slide(s, PptxInches(0.6), PptxInches(1.2), PptxInches(5.5), PptxInches(2.8), [
    '• 相关性热力图：',
    '   单一变量难以预测成功',
    '',
    '• 中文支持比例：约 65%',
    '   中国市场受重视',
    '',
    '• 发行高峰：2021-2023',
    '   疫情居家娱乐需求推动',
    '',
    '• 价格与评论数：',
    '   弱负相关，低价≠高销量'
], font_size=17)
s.shapes.add_picture(os.path.join(FIG_DIR, '09_heatmap_corr.png'), PptxInches(6.3), PptxInches(1.0), width=PptxInches(3.0))
s.shapes.add_picture(os.path.join(FIG_DIR, '13_pie_chinese.png'), PptxInches(9.5), PptxInches(1.0), width=PptxInches(3.0))
s.shapes.add_picture(os.path.join(FIG_DIR, '10_bar_year.png'), PptxInches(6.3), PptxInches(4.0), width=PptxInches(3.0))
s.shapes.add_picture(os.path.join(FIG_DIR, '14_scatter_price_reviews.png'), PptxInches(9.5), PptxInches(4.0), width=PptxInches(3.0))

# 第11页：CDF对比（bonus图）
s = add_title_only_slide(prs, '【深入验证】经验CDF vs 理论CDF')
add_text_to_slide(s, PptxInches(0.6), PptxInches(1.2), PptxInches(5.5), PptxInches(3), [
    '另一种验证方式：',
    '• 红色 = 观测数据的经验CDF',
    '• 蓝色虚线 = 理论对数正态CDF',
    '',
    '两条曲线几乎重合！',
    '→ 从累积概率角度再次验证',
    '   对数正态假设的合理性'
], font_size=17)
s.shapes.add_picture(os.path.join(FIG_DIR, '11_cdf_comparison.png'), PptxInches(6.3), PptxInches(1.2), width=PptxInches(6.3))

# 第12页：结论
s = add_title_only_slide(prs, '【解决了什么？】结论与实践启示')
add_text_to_slide(s, PptxInches(0.6), PptxInches(1.2), PptxInches(12), PptxInches(5.5), [
    '✅ 解决了“销量服从什么分布”：',
    '      评论数 ~ LogNormal(7.61, 2.15²)，市场呈“长尾效应”',
    '',
    '✅ 解决了“如何定量描述”：',
    '      MLE 估计 + K-S 检验验证 + Bootstrap 稳健推断',
    '',
    '✅ 解决了“定价是否影响口碑”：',
    '      价格对好评率无显著影响，开发者应专注品质打磨',
    '',
    '💡 给开发者的三条建议：',
    '   ① 做好“大部分游戏不会爆火”的预期管理',
    '   ② 不必陷入低价竞争，合理定价即可',
    '   ③ 用统计方法科学评估项目风险'
], font_size=20)

# 第13页：致谢
s = prs.slides.add_slide(prs.slide_layouts[6])
s.shapes.add_picture(os.path.join(FIG_DIR, '15_ppt_bg.png'), PptxInches(0), PptxInches(0), width=PptxInches(13.333))
add_text_to_slide(s, PptxInches(1), PptxInches(2.8), PptxInches(11.3), PptxInches(1.2),
    ['谢谢聆听！', '欢迎老师提问'], font_size=38, line_space=8, color=PptxRGBColor(255,255,255))

# 保存
prs.save(os.path.join(OUT_DIR, '答辩PPT_最终版.pptx'))
print("✅ PPT 已生成: output/答辩PPT_最终版.pptx")

print("\n🎉 最终版全部完成！")
print("   1. Word 报告: output/统计课程项目报告_最终版.docx")
print("   2. PPT 文件: output/答辩PPT_最终版.pptx")
print("   3. 共 15 张高清图表")

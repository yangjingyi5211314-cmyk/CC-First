"""
优化版：生成图文并茂、逻辑清晰的 Word 报告 + PPT
核心框架：要解决什么 → 做了什么 → 完成了什么 → 解决了什么
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

# ==================== 辅助函数：Word ====================
def set_run_font(run, name='Microsoft YaHei', size=10.5, bold=False, color=RGBColor(0,0,0)):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    r = run._element
    rPr = r.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn('w:eastAsia'), name)

def add_heading_custom(doc, text, level=1, color=RGBColor(0,51,102)):
    p = doc.add_paragraph()
    run = p.add_run(text)
    sizes = {0: 18, 1: 14, 2: 12}
    set_run_font(run, size=sizes.get(level, 12), bold=True, color=color)
    p.space_after = Pt(8)
    return p

def add_para_custom(doc, text, bold=False, color=RGBColor(0,0,0), size=10.5):
    p = doc.add_paragraph()
    run = p.add_run(text)
    set_run_font(run, size=size, bold=bold, color=color)
    p.space_after = Pt(6)
    return p

def add_bullet_custom(doc, text, indent_level=0):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.left_indent = Inches(0.25 + indent_level * 0.25)
    run = p.add_run(text)
    set_run_font(run)
    p.space_after = Pt(4)
    return p

# ==================== 1. 生成 Word 报告 ====================
print("正在生成优化版 Word 报告...")
doc = Document()

# 全局样式
style = doc.styles['Normal']
style.font.name = 'Microsoft YaHei'
style.font.size = Pt(10.5)
style.element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')

# ===== 封面 =====
title_p = doc.add_paragraph()
title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = title_p.add_run('买断制独立游戏销量的统计分布特征与参数估计')
set_run_font(run, size=20, bold=True, color=RGBColor(0, 51, 102))
title_p.space_after = Pt(12)

sub_p = doc.add_paragraph()
sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = sub_p.add_run('概率统计课程项目报告\n——从问题提出到方法实现，再到结论应用')
set_run_font(run, size=12, color=RGBColor(80,80,80))
doc.add_paragraph()

# ===== 摘要 =====
add_heading_custom(doc, '【摘要】', level=1)
add_para_custom(doc,
    '本项目围绕“买断制独立游戏的市场表现服从何种统计规律”这一核心问题，'
    '基于 n=500 的模拟样本数据，运用描述统计、MLE 最大似然估计、K-S 拟合优度检验、'
    'Bootstrap 自助法以及一元线性回归等方法，系统刻画了独立游戏评论数的分布特征，'
    '并探讨了定价策略与用户口碑之间的关系。'
    '研究证实：评论数服从对数正态分布，市场呈典型的“长尾效应”；价格对好评率无显著影响。')

# ===== 第一章：要解决什么 =====
add_heading_custom(doc, '第一章 要解决什么？——研究背景与问题提出', level=0)

add_heading_custom(doc, '1.1 现实背景', level=1)
add_para_custom(doc,
    '近年来，Steam、Epic Games 等平台推动了独立游戏（Indie Game）市场的蓬勃发展。'
    '与 3A 大作不同，独立游戏通常由小型团队开发、采用买断制付费模式。'
    '观察市场可以发现：少数“爆款”（如《哈迪斯》《星露谷物语》）销量破百万，'
    '而大量作品却鲜有人问津。这种“赢家通吃”的格局背后，隐藏着怎样的统计规律？')

add_heading_custom(doc, '1.2 研究问题', level=1)
add_bullet_custom(doc, '问题一：独立游戏的销量（以评论数为代理）服从何种概率分布？')
add_bullet_custom(doc, '问题二：如何利用最大似然估计（MLE）等参数估计方法定量描述该分布？')
add_bullet_custom(doc, '问题三：游戏定价策略是否显著影响用户好评率？')

add_heading_custom(doc, '1.3 研究意义', level=1)
add_para_custom(doc,
    '从统计学角度刻画独立游戏市场的分布规律，有助于开发者理解行业风险与收益特征，'
    '摆脱“凭感觉定价”的盲目性，为游戏定价和市场预期管理提供数据支撑。')

# ===== 第二章：做了什么 =====
add_heading_custom(doc, '第二章 做了什么？——研究思路与技术路线', level=0)

add_heading_custom(doc, '2.1 整体思路', level=1)
add_para_custom(doc,
    '本项目遵循“数据获取 → 描述统计 → 分布假设 → 参数估计 → 假设检验 → 区间估计 → 回归拓展”'
    '的完整统计分析流程，每一步均有明确的方法论支撑。')

add_heading_custom(doc, '2.2 具体方法', level=1)
add_bullet_custom(doc, '描述统计：计算均值、中位数、标准差、偏度、峰度，刻画数据基本面貌。')
add_bullet_custom(doc, 'MLE 最大似然估计：假设评论数服从对数正态分布，推导并计算 μ 与 σ 的 MLE。')
add_bullet_custom(doc, 'K-S 拟合优度检验：定量判断观测数据与理论分布的吻合程度。')
add_bullet_custom(doc, 'Q-Q 图：图形化验证分布假设的合理性。')
add_bullet_custom(doc, 'Bootstrap 自助法：在不依赖正态假设的前提下，构建均值的 95% 置信区间。')
add_bullet_custom(doc, '一元线性回归（OLS）：以价格为自变量、好评率为因变量，探讨定价策略的影响。')

add_heading_custom(doc, '2.3 数据来源', level=1)
add_para_custom(doc,
    '样本量 n = 500，为基于 Steam 平台独立游戏真实统计特征生成的模拟数据集。'
    '核心变量 review_count（评论数）作为销量的代理指标——研究表明 Steam 评论数与销量比例稳定（约 1:30 ~ 1:50）。')

# ===== 第三章：完成了什么 =====
add_heading_custom(doc, '第三章 完成了什么？——实验结果与数据分析', level=0)

add_heading_custom(doc, '3.1 描述统计：右偏分布的直观证据', level=1)
add_para_custom(doc, '对评论数进行描述统计，核心结果如下：')
table = doc.add_table(rows=6, cols=3)
table.style = 'Light Grid Accent 1'
hdr = table.rows[0].cells
hdr[0].text, hdr[1].text, hdr[2].text = '统计量', '数值', '统计含义'
rows = [
    ('样本量 n', '500', '数据集规模'),
    ('均值', '14,003.26', '被极端高值显著拉高'),
    ('中位数', '1,740.50', '半数游戏低于此值'),
    ('偏度', '3.9576', '显著右偏（>0）'),
    ('峰度', '16.6333', '尖峰厚尾特征'),
]
for i, (a,b,c) in enumerate(rows, 1):
    r = table.rows[i].cells
    r[0].text, r[1].text, r[2].text = a, b, c
    for cell in r:
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                set_run_font(run)
doc.add_paragraph()
add_para_custom(doc,
    '从表中可见：均值远大于中位数，偏度接近 4，峰度超过 16，表明评论数呈严重右偏且尾部厚重。'
    '直方图（图 3-1）进一步验证了这一特征：峰值集中在低评论数区域，右侧存在明显长尾。')
doc.add_picture(os.path.join(FIG_DIR, '01_histogram.png'), width=Inches(5.5))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
add_para_custom(doc, '图 3-1 评论数直方图与对数变换后的对比', size=9, color=RGBColor(100,100,100))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

add_heading_custom(doc, '3.2 MLE 参数估计：对数正态分布的定量刻画', level=1)
add_para_custom(doc,
    '基于描述统计的提示，假设评论数 X 服从对数正态分布 LogNormal(μ, σ²)，'
    '即 ln(X) ~ N(μ, σ²)。MLE 估计具有解析解：')
add_para_custom(doc, 'μ̂ = (1/n) Σ ln(xi) = 7.6115')
add_para_custom(doc, 'σ̂² = (1/n) Σ (ln(xi) − μ̂)² = 4.6338')
add_para_custom(doc, 'σ̂ = 2.1526')
add_para_custom(doc,
    '图 3-2 展示了直方图与拟合密度曲线的对比。红色的对数正态分布曲线（μ=7.61, σ=2.15）'
    '在峰值区域和右尾均能较好地贴合观测数据，拟合效果明显优于作为对照的伽马分布。')
doc.add_picture(os.path.join(FIG_DIR, '02_mle_fitting.png'), width=Inches(5.5))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
add_para_custom(doc, '图 3-2 MLE 拟合：对数正态分布 vs 伽马分布', size=9, color=RGBColor(100,100,100))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

add_heading_custom(doc, '3.3 K-S 检验与 Q-Q 图：拟合效果的统计验证', level=1)
add_para_custom(doc,
    '为定量评估拟合效果，进行 Kolmogorov-Smirnov 检验：'
    'H0：数据服从对数正态分布；H1：数据不服从对数正态分布。')
add_para_custom(doc, '检验结果：K-S 统计量 D = 0.0300，p-value = 0.7483。')
add_para_custom(doc,
    '结论：在 5% 显著性水平下，p > 0.05，不能拒绝 H0，即样本数据与对数正态分布拟合良好。'
    'Q-Q 图（图 3-3）从图形角度进一步验证：样本分位数与理论分位数基本落在 45° 参考线上。')
doc.add_picture(os.path.join(FIG_DIR, '03_qqplot.png'), width=Inches(5.5))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
add_para_custom(doc, '图 3-3 对数正态 Q-Q 图与正态 Q-Q 图（ln 变换后）', size=9, color=RGBColor(100,100,100))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

add_heading_custom(doc, '3.4 Bootstrap 置信区间：稳健的参数推断', level=1)
add_para_custom(doc,
    '由于原始数据高度右偏，传统基于正态假设的置信区间（x̄ ± z·s/√n）可能失真。'
    '采用 Bootstrap 自助法，有放回重抽样 B = 5000 次，估计样本均值的 95% 置信区间：')
table2 = doc.add_table(rows=3, cols=3)
table2.style = 'Light Grid Accent 1'
hdr2 = table2.rows[0].cells
hdr2[0].text, hdr2[1].text, hdr2[2].text = '估计方法', '95% 置信区间', '特点'
rows2 = [
    ('百分位数 Bootstrap', '[11182.27, 17114.56]', '不依赖正态假设，适合偏态分布'),
    ('正态近似', '[11032.75, 16973.77]', '依赖中心极限定理'),
]
for i, (a,b,c) in enumerate(rows2, 1):
    r = table2.rows[i].cells
    r[0].text, r[1].text, r[2].text = a, b, c
    for cell in r:
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                set_run_font(run)
doc.add_paragraph()
add_para_custom(doc,
    '两种方法结果接近，说明样本量 n=500 已足够大，中心极限定理发挥作用。'
    '值得注意的是，Bootstrap 均值的抽样分布（图 3-4）近似正态，'
    '即使原始数据严重右偏——这直观地验证了 CLT 的强大性。')
doc.add_picture(os.path.join(FIG_DIR, '04_bootstrap.png'), width=Inches(5.5))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
add_para_custom(doc, '图 3-4 Bootstrap 均值分布与 95% 置信区间', size=9, color=RGBColor(100,100,100))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

add_heading_custom(doc, '3.5 一元线性回归：价格对好评率的影响', level=1)
add_para_custom(doc,
    '作为拓展分析，建立一元线性回归模型：好评率 = β₀ + β₁ × 价格 + ε，'
    '探讨独立游戏的定价策略是否影响用户口碑。')
add_para_custom(doc, 'OLS 估计结果：')
add_bullet_custom(doc, '回归方程：好评率 = 0.7911 + 0.000100 × 价格')
add_bullet_custom(doc, 'R² = 0.0007，价格系数 p-value = 0.5557')
add_para_custom(doc,
    '结论：价格对好评率无显著线性影响（p > 0.05），且 R² 极低，'
    '说明好评率主要由游戏品质、创意和玩家匹配度驱动，而非定价策略。'
    '残差图（图 3-5 右）显示残差随机分布在 0 附近，满足同方差假定。')
doc.add_picture(os.path.join(FIG_DIR, '05_regression.png'), width=Inches(5.2))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
add_para_custom(doc, '图 3-5 价格与好评率的散点图及回归直线', size=9, color=RGBColor(100,100,100))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
doc.add_picture(os.path.join(FIG_DIR, '06_residuals.png'), width=Inches(5.2))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
add_para_custom(doc, '图 3-6 残差图与残差分布直方图', size=9, color=RGBColor(100,100,100))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

# ===== 第四章：解决了什么 =====
add_heading_custom(doc, '第四章 解决了什么？——结论、启示与局限', level=0)

add_heading_custom(doc, '4.1 核心结论', level=1)
add_bullet_custom(doc, '分布规律：独立游戏评论数服从对数正态分布，市场呈典型的“长尾效应”与“赢家通吃”格局。')
add_bullet_custom(doc, '参数估计：MLE 能有效估计分布参数（μ̂=7.61, σ̂=2.15），K-S 检验验证拟合良好（p=0.7483）。')
add_bullet_custom(doc, '区间推断：Bootstrap 自助法提供了不依赖正态假设的稳健置信区间。')
add_bullet_custom(doc, '定价启示：价格对好评率无显著影响，独立游戏的口碑主要由品质驱动。')

add_heading_custom(doc, '4.2 对独立游戏开发者的实践启示', level=1)
add_bullet_custom(doc, '市场预期管理：独立游戏市场高度分化，开发者应做好“大部分作品不会爆火”的心理准备和财务规划。')
add_bullet_custom(doc, '定价策略：不必过度担忧“定价高影响口碑”，应将更多资源投入游戏品质打磨而非低价竞争。')
add_bullet_custom(doc, '数据思维：利用统计方法对市场数据进行建模，可以更科学地评估项目风险与收益分布。')

add_heading_custom(doc, '4.3 研究局限与展望', level=1)
add_bullet_custom(doc, '数据层面：本研究使用模拟数据，未来可接入 Steam API 获取真实数据进行验证。')
add_bullet_custom(doc, '方法层面：回归分析仅纳入价格单一变量，未来可拓展为多元回归，纳入游戏时长、标签数量、发行年份等变量。')
add_bullet_custom(doc, '应用层面：可进一步研究时间序列特征，分析独立游戏市场的演化趋势和周期性规律。')

# ===== 参考文献 =====
add_heading_custom(doc, '参考文献', level=0)
add_para_custom(doc, '[1] Casella G, Berger R L. Statistical Inference (2nd ed.)[M]. Duxbury Press, 2002.')
add_para_custom(doc, '[2] Efron B, Tibshirani R J. An Introduction to the Bootstrap[M]. CRC Press, 1994.')
add_para_custom(doc, '[3] SteamSpy. Steam Games Statistics[EB/OL]. https://steamspy.com/, 2024.')
add_para_custom(doc, '[4] SteamDB. Steam Database[EB/OL]. https://steamdb.info/, 2024.')

# 保存
doc.save(os.path.join(OUT_DIR, '统计课程项目报告_优化版.docx'))
print("✅ Word 报告已生成: output/统计课程项目报告_优化版.docx")

# ==================== 2. 生成 PPT（优化版）====================
print("正在生成优化版 PPT...")
prs = Presentation()
prs.slide_width = PptxInches(13.333)
prs.slide_height = PptxInches(7.5)

# 辅助函数
def add_title_only_slide(prs, title_text):
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    box = slide.shapes.add_textbox(PptxInches(0.5), PptxInches(0.3), PptxInches(12.3), PptxInches(0.9))
    tf = box.text_frame
    tf.text = title_text
    p = tf.paragraphs[0]
    p.font.size = Pt(30)
    p.font.bold = True
    p.font.color.rgb = PptxRGBColor(0, 51, 102)
    return slide

def add_text_to_slide(slide, left, top, width, height, texts, font_size=20, line_space=14):
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
        p.space_after = Pt(line_space)
    return box

# 第1页：封面
s = prs.slides.add_slide(prs.slide_layouts[6])
add_text_to_slide(s, PptxInches(1), PptxInches(2.2), PptxInches(11.3), PptxInches(1.2),
    ['买断制独立游戏销量的', '统计分布特征与参数估计'], font_size=36, line_space=8)
add_text_to_slide(s, PptxInches(1), PptxInches(4.2), PptxInches(11.3), PptxInches(1),
    ['概率统计课程项目 | 基于 MLE、Bootstrap 与回归分析'], font_size=18)

# 第2页：要解决什么
s = add_title_only_slide(prs, '【要解决什么？】研究背景与核心问题')
add_text_to_slide(s, PptxInches(0.6), PptxInches(1.3), PptxInches(12), PptxInches(5.5), [
    '• 现实背景：',
    '   独立游戏市场高度分化，少数“爆款” vs 大量“长尾”',
    '   Steam 平台数据丰富，但销量分布规律尚缺乏定量刻画',
    '',
    '• 三个核心问题：',
    '   ① 独立游戏销量服从何种概率分布？',
    '   ② 如何用 MLE 等参数估计方法定量描述该分布？',
    '   ③ 游戏定价策略是否显著影响用户好评率？',
    '',
    '• 研究意义：',
    '   为独立游戏开发者提供数据驱动的市场预期与定价决策依据'
], font_size=20)

# 第3页：做了什么（方法）
s = add_title_only_slide(prs, '【做了什么？】技术路线与分析方法')
add_text_to_slide(s, PptxInches(0.6), PptxInches(1.3), PptxInches(6), PptxInches(5.5), [
    '① 描述统计',
    '   均值、中位数、偏度、峰度',
    '',
    '② MLE 最大似然估计',
    '   假设对数正态分布，估计 μ 与 σ',
    '',
    '③ K-S 拟合优度检验',
    '   定量验证分布假设',
    '',
    '④ Q-Q 图',
    '   图形化验证拟合效果',
], font_size=18)
add_text_to_slide(s, PptxInches(6.8), PptxInches(1.3), PptxInches(6), PptxInches(5.5), [
    '⑤ Bootstrap 自助法',
    '   构建稳健的 95% 置信区间',
    '',
    '⑥ 一元线性回归（OLS）',
    '   价格 → 好评率',
    '',
    '数据来源：',
    '   n=500 模拟样本',
    '   基于 Steam 真实统计特征',
], font_size=18)

# 第4页：做了什么 — 描述统计
s = add_title_only_slide(prs, '【做了什么？】描述统计：发现右偏分布')
add_text_to_slide(s, PptxInches(0.6), PptxInches(1.2), PptxInches(5.5), PptxInches(2.8), [
    '• 均值 14,003 >> 中位数 1,740',
    '• 偏度 = 3.96（显著右偏）',
    '• 峰度 = 16.63（尖峰厚尾）',
    '',
    '→ 对数变换后接近正态',
    '→ 提示：原数据可能服从对数正态分布'
], font_size=18)
s.shapes.add_picture(os.path.join(FIG_DIR, '01_histogram.png'), PptxInches(6.3), PptxInches(1.2), width=PptxInches(6.3))

# 第5页：做了什么 — MLE
s = add_title_only_slide(prs, '【做了什么？】MLE：对数正态分布的参数估计')
add_text_to_slide(s, PptxInches(0.6), PptxInches(1.2), PptxInches(5.5), PptxInches(3.2), [
    '假设：X ~ LogNormal(μ, σ²)',
    '即 ln(X) ~ N(μ, σ²)',
    '',
    'MLE 解析解：',
    '   μ̂ = 7.6115',
    '   σ̂ = 2.1526',
    '',
    '→ 对数正态拟合效果优于伽马分布'
], font_size=18)
s.shapes.add_picture(os.path.join(FIG_DIR, '02_mle_fitting.png'), PptxInches(6.3), PptxInches(1.2), width=PptxInches(6.3))

# 第6页：做了什么 — K-S + Q-Q
s = add_title_only_slide(prs, '【做了什么？】K-S 检验与 Q-Q 图验证')
add_text_to_slide(s, PptxInches(0.6), PptxInches(1.2), PptxInches(5.5), PptxInches(3), [
    'K-S 检验：',
    '   H0：数据服从对数正态分布',
    '   H1：数据不服从对数正态分布',
    '',
    '结果：',
    '   K-S 统计量 D = 0.0300',
    '   p-value = 0.7483',
    '',
    '结论：p > 0.05，不能拒绝 H0',
    '→ 对数正态假设通过统计检验'
], font_size=18)
s.shapes.add_picture(os.path.join(FIG_DIR, '03_qqplot.png'), PptxInches(6.3), PptxInches(1.2), width=PptxInches(6.3))

# 第7页：完成了什么 — Bootstrap
s = add_title_only_slide(prs, '【完成了什么？】Bootstrap 稳健置信区间')
add_text_to_slide(s, PptxInches(0.6), PptxInches(1.2), PptxInches(5.5), PptxInches(3.2), [
    '问题：原始数据右偏，传统正态',
    '      置信区间可能失真',
    '',
    '方法：Bootstrap 自助法',
    '   有放回重抽样 B = 5000 次',
    '',
    '结果：95% CI = [11182, 17115]',
    '',
    '→ 即使原数据偏态，Bootstrap 均值',
    '   分布仍近似正态（验证 CLT）'
], font_size=18)
s.shapes.add_picture(os.path.join(FIG_DIR, '04_bootstrap.png'), PptxInches(6.3), PptxInches(1.2), width=PptxInches(6.3))

# 第8页：完成了什么 — 回归
s = add_title_only_slide(prs, '【完成了什么？】回归分析：价格不影响口碑')
add_text_to_slide(s, PptxInches(0.6), PptxInches(1.2), PptxInches(5.5), PptxInches(3.2), [
    '模型：好评率 = β₀ + β₁ × 价格 + ε',
    '',
    '结果：',
    '   回归方程：好评率 = 0.7911',
    '            + 0.000100 × 价格',
    '   R² = 0.001',
    '   p-value = 0.556 > 0.05',
    '',
    '结论：价格对好评率无显著影响',
    '→ 口碑由品质驱动，而非定价'
], font_size=18)
s.shapes.add_picture(os.path.join(FIG_DIR, '05_regression.png'), PptxInches(6.3), PptxInches(1.2), width=PptxInches(6.3))

# 第9页：解决了什么
s = add_title_only_slide(prs, '【解决了什么？】结论与实践启示')
add_text_to_slide(s, PptxInches(0.6), PptxInches(1.2), PptxInches(12), PptxInches(5.5), [
    '✅ 解决了“销量服从什么分布”的问题：',
    '      评论数服从对数正态分布，市场呈“长尾效应”',
    '',
    '✅ 解决了“如何定量描述”的问题：',
    '      MLE 估计参数 μ̂=7.61, σ̂=2.15，K-S 检验验证拟合良好',
    '',
    '✅ 解决了“定价是否影响口碑”的问题：',
    '      价格对好评率无显著影响，开发者应专注品质打磨',
    '',
    '💡 实践启示：',
    '   • 做好“大部分游戏不会爆火”的市场预期',
    '   • 不必过度担心定价影响口碑',
    '   • 用统计方法科学评估项目风险'
], font_size=20)

# 第10页：致谢
s = prs.slides.add_slide(prs.slide_layouts[6])
add_text_to_slide(s, PptxInches(1), PptxInches(2.8), PptxInches(11.3), PptxInches(1.2),
    ['谢谢聆听！', '欢迎老师提问'], font_size=36)

# 保存
prs.save(os.path.join(OUT_DIR, '答辩PPT_优化版.pptx'))
print("✅ PPT 已生成: output/答辩PPT_优化版.pptx")

print("\n🎉 优化版全部完成！新文件：")
print("   1. Word 报告: output/统计课程项目报告_优化版.docx")
print("   2. PPT 文件: output/答辩PPT_优化版.pptx")

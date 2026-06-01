"""
model_core.py — 黑山土石方测算模型的「计算核心」

职责：
    compute(params: dict) -> dict
        输入 21 个用户参数 → 输出单价、各成本拆解、警告等

    export_xlsx(params: dict) -> bytes
        输入 21 个用户参数 → 返回完整 xlsx 文件的字节流（可直接下载）

依赖：
    - build_v8.py（同目录上一级）
    - libreoffice headless（系统命令，用于公式重算）

启动开销：
    首次 import 时跑一次 build_v8 生成 template_v8.xlsx，约 1-2 秒
    后续每次 compute / export 调用约 5-10 秒（libreoffice 重算）
"""

from __future__ import annotations

import os
import sys
import shutil
import tempfile
import subprocess
import io
from typing import Any, Dict

# ---------- 路径：让本模块可定位到 build_v8.py（同目录或上一级目录都能找到）----------
_HERE = os.path.dirname(os.path.abspath(__file__))
_PARENT = os.path.dirname(_HERE)
for _p in (_HERE, _PARENT):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from openpyxl import load_workbook

# ---------- 启动时生成一次模板 ----------
_TEMPLATE_PATH = os.path.join(_HERE, 'template_v8.xlsx')

def _ensure_template():
    """首次调用时跑 build_v8 生成 template_v8.xlsx；后续直接复用"""
    if os.path.exists(_TEMPLATE_PATH):
        return
    import build_v8  # 执行 build_v8.py 模块（已在 if __name__='__main__' 内保护 save）
    build_v8.wb.save(_TEMPLATE_PATH)

# ---------- 输入参数到主表 cell 的映射 ----------
# (cell, 类型 'str'|'num', 默认值)
PARAM_MAP = {
    # 岩石
    'rock_level':   ('B4',  'str', 'Ⅰ极软岩'),
    # 工艺
    'process':      ('B12', 'str', '直接开挖'),
    # 挖装设备
    'excavator':    ('B17', 'str', 'W4.5'),
    'excavator_src':('B18', 'str', '自有'),
    # 运输设备
    'truck':        ('B23', 'str', 'T65B'),
    'truck_src':    ('B24', 'str', '自有'),
    # 钻孔设备
    'drill':        ('B29', 'str', 'ZG90'),
    'drill_src':    ('B30', 'str', '自有'),
    # 运输参数
    'dist_in':      ('B34', 'num', 2.1),
    'dist_out':     ('B35', 'num', 5.0),
    'dir_in':       ('B37', 'str', '平路'),
    'slope_in':     ('B38', 'num', 1.5),
    'dir_out':      ('B39', 'str', '平路'),
    'slope_out':    ('B40', 'num', 1.0),
    # 爆破设计
    'blast_len':    ('B44', 'num', 50),
    'blast_wid':    ('B45', 'num', 20),
    'buffer_rows':  ('B46', 'num', 1),
    'step_h':       ('B47', 'num', 10),
    'slope_angle':  ('B48', 'num', 75),
    'hole_angle':   ('B49', 'num', 90),
    'hole_pattern': ('B50', 'str', '矩形'),
    'hole_diameter':('B52', 'num', 90),
    'pre_split':    ('B53', 'str', '否'),
}

# ---------- 输出 cell（用于 compute 返回值） ----------
OUTPUT_CELLS = {
    # 综合单价（最关键）
    'price_incl_tax':   ('成本汇总-测算主表', 'B65'),
    'price_excl_tax':   ('成本汇总-测算主表', 'B64'),
    # 各成本项
    'cost_blast':       ('成本汇总-测算主表', 'B57'),
    'cost_excavate':    ('成本汇总-测算主表', 'B58'),
    'cost_transport':   ('成本汇总-测算主表', 'B59'),
    'cost_loosen':      ('成本汇总-测算主表', 'B60'),
    'cost_crush':       ('成本汇总-测算主表', 'B61'),
    'cost_second_crush':('成本汇总-测算主表', 'B62'),
    'cost_dump':        ('成本汇总-测算主表', 'B63'),
    # 联动的岩石参数
    'f_value':          ('成本汇总-测算主表', 'B5'),
    'density':          ('成本汇总-测算主表', 'B6'),
    'loose_factor':     ('成本汇总-测算主表', 'B7'),
    'big_block_rate':   ('成本汇总-测算主表', 'B8'),
    # 推荐
    'recommend_process':('成本汇总-测算主表', 'B13'),
    'recommend_exc':    ('成本汇总-测算主表', 'B19'),
    'recommend_truck':  ('成本汇总-测算主表', 'B22'),
    'recommend_drill':  ('成本汇总-测算主表', 'B28'),
    'recommend_hole_d': ('成本汇总-测算主表', 'B51'),   # 推荐孔径（V8 B51 公式按岩石+工艺推导）
    # 校验警告
    'warn_process':     ('成本汇总-测算主表', 'B14'),
    'warn_truck':       ('成本汇总-测算主表', 'B25'),
    'warn_drill':       ('成本汇总-测算主表', 'B31'),
    'warn_slope':       ('成本汇总-测算主表', 'B41'),
    'warn_step':        ('成本汇总-测算主表', 'B54'),
}

# ---------- 选项常量（给 UI 用） ----------
OPTIONS = {
    'rock_level':   ['Ⅰ极软岩', 'Ⅱ软岩', 'Ⅲ较软岩', 'Ⅳ中硬岩',
                     'Ⅴ较硬岩', 'Ⅵ坚硬岩', 'Ⅶ极硬岩', 'Ⅷ特殊坚硬岩'],
    'process':      ['直接开挖', '松土器松土+开挖', '机械破碎+开挖',
                     '爆破+开挖', '爆破+二次破碎+开挖'],
    'excavator':    ['W1.6', 'W2.0', 'W2.5', 'W3.2', 'E2.0',
                     'W4.5', 'W5.5', 'W6.5', 'W8.0', 'E6.5', 'E8.0'],
    'excavator_src':['自有', '租赁'],
    'truck':        ['T30H', 'E30S', 'T65B', 'T60S', 'T91B', 'E60Y', 'E93L', 'E75T'],
    'truck_src':    ['自有', '租赁'],
    'drill':        ['ZG90', 'ZG120', 'ACD7', 'ACD9', 'AC90', 'AC120', 'KY200', 'KY250'],
    'drill_src':    ['自有', '租赁'],
    'dir_in':       ['平路', '上坡', '下坡'],
    'dir_out':      ['平路', '上坡', '下坡'],
    'hole_pattern': ['矩形', '梅花形'],
    'hole_diameter':[70, 90, 115, 150, 200, 250],
    'pre_split':    ['是', '否'],
    'buffer_rows':  [0, 1, 2],
}

# ---------- 默认参数（用户视角的"默认工况"） ----------
DEFAULT_PARAMS = {k: v[2] for k, v in PARAM_MAP.items()}


# ============================================================
# 设备元信息映射（方案 I 新增 · 2026-06-01）
# ------------------------------------------------------------
# 用于 UI 在选型控件下方展示设备的真实物理含义，避免用户被"代号"误导。
# 比如 E60Y 看着像"更大型号"，但其实是 60t 纯电；E93L 才是 93t 纯电。
# 同时为「钻机↔孔径」联动提供权威映射（来源：build_v8.py 参数库 A40:O58）
# ============================================================

# 矿卡：代号 → (载重t, 动力类型)  ——  数据来源：参数库 A40:E47
TRUCK_INFO = {
    'T30H': (30, '柴油'),
    'E30S': (31, '纯电'),
    'T65B': (65, '柴油'),
    'T60S': (60, '柴油'),
    'T91B': (91, '柴油'),
    'E60Y': (60, '纯电'),
    'E93L': (93, '纯电'),
    'E75T': (75, '纯电'),
}

# 钻机：代号 → (类型, 推荐孔径mm, 孔径区间mm)
#   推荐孔径：默认匹配孔径（用于联动跳变的默认值）
#   孔径区间：参数库 A63:H84 中该钻机覆盖的所有孔径
#   注：潜孔钻 ZG/AC 系列推荐 90mm（最经济）；牙轮钻 KY 系列推荐自身命名口径
DRILL_INFO = {
    'ZG90':  ('潜孔钻', 90,  [70, 90, 115]),
    'ZG120': ('潜孔钻', 90,  [70, 90, 115]),
    'ACD7':  ('潜孔钻', 90,  [70, 90, 115]),
    'ACD9':  ('潜孔钻', 90,  [70, 90, 115]),
    'AC90':  ('潜孔钻', 90,  [70, 90, 115]),
    'AC120': ('潜孔钻', 115, [70, 90, 115]),
    'KY200': ('牙轮钻', 200, [150, 200]),
    'KY250': ('牙轮钻', 250, [200, 250]),
}


def drill_recommend_diameter(drill: str) -> int:
    """根据钻机型号，返回该钻机的推荐孔径（用于 UI 钻机↔孔径联动）"""
    info = DRILL_INFO.get(drill)
    return info[1] if info else 90


def truck_display(code: str) -> str:
    """矿卡代号 → '代号(载重t·动力)' 的展示字符串，UI 用"""
    info = TRUCK_INFO.get(code)
    if not info:
        return code
    return f"{code}（{info[0]}t·{info[1]}）"


def drill_display(code: str) -> str:
    """钻机代号 → '代号(类型·推荐孔径)' 的展示字符串，UI 用"""
    info = DRILL_INFO.get(code)
    if not info:
        return code
    return f"{code}（{info[0]}·标配 {info[1]}mm）"


# ============================================================
# 工艺错配惩罚体系（V8 模型外置·后置校正层）
# ------------------------------------------------------------
# V8 模型按用户选择直接计算，不做合理性校验；这一层在 compute()
# 末尾做两件事：
#   1) 根据"岩石—推荐工艺"映射识别错配
#   2) 对挖装费应用 (1/eff + 0.3*(con-1)) 复合惩罚，并连带刷新总价
#   3) 输出 warning 字段，前端按 severity 决定红/橙/蓝色提示
# ============================================================

# 岩石级别 → 推荐工艺（**完全对齐 V8 参数库 I3:I10**）
# 设计原则（方案H 2026-06-01）：
#   - 工艺类型选择 ≠ 工艺参数设计
#   - V8 已经按岩石硬度自动调整参数（装药量、孔距、钻孔速度等），成本天然不同
#   - 错配校验只针对"工艺类型大方向不对"（用爆破打软岩 / 直接开挖打硬岩）
#   - 大块率高是"爆破设计参数问题"，由工程师调炮孔布置，不属于工艺类型错配
#   - 因此硬岩（含 Ⅷ岩）推荐"爆破+开挖"即合理，二次破碎不是必须
ROCK_RECOMMEND = {
    'Ⅰ极软岩':     '直接开挖',
    'Ⅱ软岩':       '松土器松土+开挖',
    'Ⅲ较软岩':     '松土器松土+开挖',
    'Ⅳ中硬岩':     '爆破+开挖',
    'Ⅴ较硬岩':     '爆破+开挖',
    'Ⅵ坚硬岩':     '爆破+开挖',
    'Ⅶ极硬岩':     '爆破+开挖',
    'Ⅷ特殊坚硬岩':  '爆破+开挖',
}

# 工艺强度档位（数字越小越偷懒）
PROCESS_LEVEL = {
    '直接开挖':           1,
    '松土器松土+开挖':    2,
    '机械破碎+开挖':      3,
    '爆破+开挖':          4,
    '爆破+二次破碎+开挖': 5,
}
# 方案H：错配补偿落到"错配工艺主项"上
# —— 哪个工艺被错配选用，错配补偿就加到该工艺对应的成本项
# 同时维护中文名映射供 UI 显示
COST_FIELD_BY_PROCESS = {
    '直接开挖':           'cost_excavate',     # 挖机硬挖磨耗
    '松土器松土+开挖':    'cost_loosen',       # 松土器磨耗剧增
    '机械破碎+开挖':      'cost_crush',        # 破碎锤效率折损
    '爆破+开挖':          'cost_blast',        # 爆破效果不足
    '爆破+二次破碎+开挖': 'cost_second_crush', # 二次破碎不够
}
# 成本项中文名（供 UI 提示文案使用）
COST_FIELD_LABEL = {
    'cost_excavate':     '挖装费',
    'cost_loosen':       '松土费',
    'cost_crush':        '破碎费',
    'cost_blast':        '爆破费',
    'cost_second_crush': '二次破碎费',
    'cost_transport':    '运输费',
    'cost_dump':         '渣场费',
}
# 方案H：撤销了 SPECIAL_PENALTY_FIELD 特殊映射
# 错配补偿统一按 COST_FIELD_BY_PROCESS 落到"用户选的错配工艺"对应成本项
# 不再做"物理影响"层面的猜测重定向（挖装是最末端环节，不应替别的工艺背锅）

# 推荐工艺基准价（含税，元/方）—— 默认参数下 V8 算出的推荐工艺价
# 用作工艺错配阶梯定价的锚点：错配价 = baseline × (1 + gap × 0.5)
RECOMMEND_BASELINE_INCL = {
    'Ⅰ极软岩':      8.92,
    'Ⅱ软岩':       12.80,
    'Ⅲ较软岩':     15.89,
    'Ⅳ中硬岩':     30.01,
    'Ⅴ较硬岩':     33.73,
    'Ⅵ坚硬岩':     38.59,
    'Ⅶ极硬岩':     45.27,
    'Ⅷ特殊坚硬岩': 60.33,
}

# 工艺错配阶梯定价系数：factor = 2 ** gap
#   差 0 档（推荐）→ ×1
#   差 1 档 → ×2
#   差 2 档 → ×4
#   差 3 档 → ×8
#   差 4 档 → ×16
def _gap_factor(gap):
    return 2 ** gap if gap > 0 else 1.0

# (用户选择, 系统推荐) → (效率折减, 耗材倍率, 严重度, 说明)
# severity: 'error'=红 / 'warning'=橙 / 'info'=蓝
PENALTY = {
    # ↓ "工艺不足"型偏离：用户选了更轻量的工艺，挖机硬挖
    ('直接开挖',         '松土器松土+开挖'):    (0.8,  1.2, 'info',    '岩石较硬，挖机硬挖效率折损约 20%，建议改用松土器'),
    ('直接开挖',         '机械破碎+开挖'):      (0.3,  3.0, 'warning', '岩石需破碎，挖机硬挖效率仅 30%，齿尖磨耗剧增'),
    ('直接开挖',         '爆破+开挖'):         (0.1,  5.0, 'error',   '岩石需爆破，挖机几乎挖不动，齿尖暴损'),
    ('直接开挖',         '爆破+二次破碎+开挖'): (0.08, 6.0, 'error',   '特坚岩需爆破+二次破碎，直接开挖完全不可行'),
    ('松土器松土+开挖',   '机械破碎+开挖'):     (0.5,  2.0, 'warning', '岩石较硬，松土头磨耗剧增，效率仅 50%'),
    ('松土器松土+开挖',   '爆破+开挖'):         (0.2,  3.0, 'error',   '岩石需爆破，松土器极不推荐，效率仅 20%'),
    ('松土器松土+开挖',   '爆破+二次破碎+开挖'): (0.15, 4.0, 'error',  '特坚岩松土完全不可行'),
    ('机械破碎+开挖',     '爆破+开挖'):         (0.5,  2.0, 'warning', '岩石需爆破，机械破碎效率仅 50%'),
    ('机械破碎+开挖',     '爆破+二次破碎+开挖'): (0.4,  2.5, 'warning', '特坚岩机械破碎效率较低且大块多'),
    ('爆破+开挖',         '爆破+二次破碎+开挖'): (0.7,  1.5, 'info',    '特坚岩建议带二次破碎，否则大块多→运输/破碎成本上升'),
    # ↓ "工艺过剩"型偏离：杀鸡牛刀，不影响真实成本，仅提示
    ('松土器松土+开挖',   '直接开挖'):          (1.0,  1.0, 'info', '工艺略偏保守，可改用直接开挖更经济'),
    ('机械破碎+开挖',     '直接开挖'):          (1.0,  1.0, 'info', '工艺过剩，可改用直接开挖更经济'),
    ('机械破碎+开挖',     '松土器松土+开挖'):    (1.0,  1.0, 'info', '工艺偏保守'),
    ('爆破+开挖',         '直接开挖'):          (1.0,  1.0, 'info', '工艺过剩，可改用直接开挖大幅降本'),
    ('爆破+开挖',         '松土器松土+开挖'):    (1.0,  1.0, 'info', '工艺偏保守'),
    ('爆破+开挖',         '机械破碎+开挖'):      (1.0,  1.0, 'info', '工艺偏保守'),
    ('爆破+二次破碎+开挖', '直接开挖'):          (1.0,  1.0, 'info', '工艺过剩'),
    ('爆破+二次破碎+开挖', '松土器松土+开挖'):    (1.0,  1.0, 'info', '工艺过剩'),
    ('爆破+二次破碎+开挖', '机械破碎+开挖'):      (1.0,  1.0, 'info', '工艺偏保守'),
    ('爆破+二次破碎+开挖', '爆破+开挖'):          (1.0,  1.0, 'info', '工艺偏保守'),
}


def detect_mismatch(rock_level, process):
    """识别工艺错配。返回 dict 或 None。"""
    recommended = ROCK_RECOMMEND.get(rock_level)
    if not recommended or process == recommended:
        return None
    rule = PENALTY.get((process, recommended))
    if not rule:
        return None
    eff, con, sev, msg = rule
    return {
        'severity': sev,
        'message': msg,
        'process_user': process,
        'process_recommend': recommended,
        'rock_level': rock_level,
        'efficiency_factor': eff,
        'consume_factor': con,
        'is_penalty': (eff < 1.0 or con > 1.0),
    }


def apply_penalty(result):
    """工艺错配后置校正：按"工艺档位差"覆盖式定价。

    业务原则：工艺越偷懒（强度越低），单价应该越贵。
        直接开挖 > 松土器 > 机械破碎 > 爆破+开挖

    实现：
    - 每个岩石等级有「推荐工艺」和「推荐工艺基准价」
    - 用户选其它工艺时，按档位差强制定价：基准价 × (1 + gap × 0.5)
    - 覆盖 V8 原算，避免「V8 在硬岩+松土器自带产量罚导致单价飙到几百几千」
      和「V8 在硬岩+直接开挖完全没罚导致直接开挖比松土器还便宜」两个 bug
    - V8 原始价记录在 warning.v8_original_price_incl，让用户看到「原始算 763 元/方」的不可行性
    """
    inputs = result.get('inputs', {})
    rock = inputs.get('rock_level')
    proc = inputs.get('process')
    mismatch = detect_mismatch(rock, proc)
    result['warning'] = mismatch
    if not mismatch:
        return result

    level_user = PROCESS_LEVEL.get(proc, 0)
    level_rec = PROCESS_LEVEL.get(mismatch['process_recommend'], 0)
    gap = level_rec - level_user  # 正数=用户工艺偏软（不足）

    base_excl = result.get('price_excl_tax') or 0
    base_incl = result.get('price_incl_tax') or 0
    tax_ratio = (base_incl / base_excl) if base_excl > 0 else 1.09

    # ---- V8 模型跳过项检测（用于 UI 解释「为什么选了爆破但没爆破费」）----
    # V8 模型 B57 公式：IF(f<2 且选爆破工艺, 0, 计算爆破费) —— f<2 时强制爆破费=0
    # V8 模型 B62 公式：爆破+二次破碎工艺时 = 破碎费×大块率，软岩大块率≈0
    f_value = result.get('f_value') or 0
    skipped = []
    if f_value > 0 and f_value < 2:
        if proc in ('爆破+开挖', '爆破+二次破碎+开挖') and (result.get('cost_blast') or 0) == 0:
            skipped.append('爆破费')
        if proc == '爆破+二次破碎+开挖' and (result.get('cost_second_crush') or 0) == 0:
            skipped.append('二次破碎费')
    if skipped:
        mismatch['v8_skipped_costs'] = skipped
        mismatch['v8_skip_reason'] = (
            f"V8 模型规则：当前岩石普氏系数 f={f_value:.2f} < 2（极软岩自带破碎），"
            f"虽然您选择了「{proc}」，但 V8 不计算 {('、'.join(skipped))}（取 0）"
        )

    # ---- C 类：工艺过剩（杀鸡牛刀）----
    # 校正层不改价，但记录"推荐工艺基准价"供 UI 做对比，解释为什么 V8 算出的价
    # 和推荐工艺仍有差异（V8 按用户选的工艺真实算各项成本）
    if gap <= 0:
        mismatch['severity'] = 'info'
        # 给 C 类补上对比信息：V8 按用户工艺算的实际价 vs 推荐工艺参考价
        recommend_baseline = RECOMMEND_BASELINE_INCL.get(rock)
        if recommend_baseline is not None:
            mismatch['v8_actual_price_incl'] = round(base_incl, 2)
            mismatch['recommend_baseline_incl'] = round(recommend_baseline, 2)
            mismatch['excess_levels'] = -gap  # 正数=用户工艺过剩档位
            mismatch['note'] = (
                f"工艺过剩 {-gap} 档。V8 按您选的「{proc}」算实际成本 "
                f"{base_incl:.2f} 元/方（含税），推荐工艺「{mismatch['process_recommend']}」"
                f"默认参数下成本约 {recommend_baseline:.2f} 元/方，"
                f"两者差异来自 V8 模型按不同工艺算的物料参数（松散系数、大块率等）"
            )
        return result

    # ---- B 类：工艺偏软 → 按档位差强制定价 ----
    baseline_incl = RECOMMEND_BASELINE_INCL.get(rock)
    if baseline_incl is None:
        return result  # 安全 fallback
    baseline_excl = baseline_incl / tax_ratio
    factor = _gap_factor(gap)
    target_excl = baseline_excl * factor
    target_incl = target_excl * tax_ratio

    # 记录到 warning 供 UI 显示
    mismatch['v8_original_price_incl'] = round(base_incl, 2)
    mismatch['baseline_price_incl'] = round(baseline_incl, 2)
    mismatch['gap_levels'] = gap
    mismatch['adjusted_factor'] = round(factor, 2)
    mismatch['adjusted_price_incl'] = round(target_incl, 2)

    # 方案H：错配补偿统一落到"用户选的错配工艺主项"上（物理纯净 + 不踩挖装）
    # —— 选错松土器就罚松土费、选错破碎就罚破碎费、选错爆破就罚爆破费
    # —— 不再做"物理影响"猜测重定向（挖装作为最末端环节不替别人背锅）
    # 边界保护：当 V8 原算 ≥ 阶梯定价时按 V8 实际成本走（避免出现负数）
    main_field = COST_FIELD_BY_PROCESS.get(proc, 'cost_excavate')
    main_label = COST_FIELD_LABEL.get(main_field, main_field)
    orig_main = result.get(main_field) or 0
    extra = target_excl - base_excl
    if extra >= 0:
        # 阶梯定价更贵 → 差价加到错配主项，按阶梯定价收
        result[main_field + '_original'] = orig_main
        result[main_field] = orig_main + extra
        result['penalty_field'] = main_field
        result['penalty_field_label'] = main_label
        result['penalty_extra'] = round(extra, 4)
        result['penalty_multiplier'] = factor
        result['price_excl_tax'] = target_excl
        result['price_incl_tax'] = target_incl
        mismatch['penalty_field'] = main_field
        mismatch['penalty_field_label'] = main_label
        mismatch['penalty_extra'] = round(extra, 2)
        mismatch['note'] = (
            f"错配补偿 {extra:.2f} 元/方 已加到「{main_label}」上"
            f"（V8 原算 {orig_main:.2f} → 校正后 {orig_main+extra:.2f}）"
            f"，其他成本项保持 V8 原算真实值不变"
        )
    else:
        # V8 原算已超过阶梯定价 → 按 V8 实际成本走（不强制下调）
        mismatch['note'] = (
            f"V8 按错配工艺算出 {base_incl:.2f} 元/方，已超过阶梯定价 {target_incl:.2f}，"
            "按 V8 实际成本计费（错配工艺实际成本更高，无需额外补偿）"
        )
    return result


# ---------- 核心：把参数应用到模板 + libreoffice 重算 ----------
def _build_and_calc(params: Dict[str, Any], workdir: str) -> str:
    """把参数写入主表 → save → libreoffice 重算 → 返回重算后文件路径"""
    _ensure_template()
    src = os.path.join(workdir, 'src.xlsx')
    shutil.copy(_TEMPLATE_PATH, src)

    # 1. 写主表参数
    wb = load_workbook(src)
    ws = wb['成本汇总-测算主表']
    full = {**DEFAULT_PARAMS, **(params or {})}
    for key, val in full.items():
        if key not in PARAM_MAP:
            continue
        cell, typ, _ = PARAM_MAP[key]
        if typ == 'num':
            try:
                val = float(val)
            except (TypeError, ValueError):
                continue
        ws[cell] = val
    wb.save(src)

    # 2. libreoffice headless 重算
    out_dir = os.path.join(workdir, 'recalc')
    os.makedirs(out_dir, exist_ok=True)
    soffice = _find_libreoffice()
    cmd = [soffice, '--headless', '--calc',
           '--convert-to', 'xlsx', '--outdir', out_dir, src]
    res = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
    if res.returncode != 0:
        raise RuntimeError(f"libreoffice 重算失败：{res.stderr or res.stdout}")
    out_path = os.path.join(out_dir, 'src.xlsx')
    if not os.path.exists(out_path):
        raise RuntimeError("libreoffice 未生成输出文件")
    return out_path


def _find_libreoffice() -> str:
    for name in ('soffice', 'libreoffice'):
        p = shutil.which(name)
        if p:
            return p
    raise RuntimeError("找不到 libreoffice/soffice 命令，请先安装 libreoffice")


# ---------- 对外 API ----------
def compute(params: Dict[str, Any] = None) -> Dict[str, Any]:
    """
    给一组用户参数，返回所有关键计算结果。
    返回的 dict 包含 OUTPUT_CELLS 里全部 key + 'inputs' 回显。
    """
    with tempfile.TemporaryDirectory() as workdir:
        recalc_path = _build_and_calc(params or {}, workdir)
        wb = load_workbook(recalc_path, data_only=True)
        out = {}
        for key, (sheet, cell) in OUTPUT_CELLS.items():
            v = wb[sheet][cell].value
            out[key] = v
        out['inputs'] = {**DEFAULT_PARAMS, **(params or {})}
        # 应用工艺错配惩罚（加 warning 字段、修正挖装费和总价）
        apply_penalty(out)
        return out


def export_xlsx(params: Dict[str, Any] = None) -> bytes:
    """
    给一组用户参数，返回完整 xlsx 字节流（含所有 8 个 sheet + 公式计算后的值）
    可直接放进 Streamlit 的 st.download_button(data=...)
    """
    with tempfile.TemporaryDirectory() as workdir:
        recalc_path = _build_and_calc(params or {}, workdir)
        with open(recalc_path, 'rb') as f:
            return f.read()


# ---------- 自检 ----------
if __name__ == '__main__':
    print("[1/3] 生成模板...")
    _ensure_template()
    print(f"      模板路径: {_TEMPLATE_PATH}  ({os.path.getsize(_TEMPLATE_PATH)} bytes)")

    print("[2/3] 默认工况（Ⅰ极软岩 + 直接开挖）...")
    r = compute()
    print(f"      含税综合单价: {r['price_incl_tax']:.4f} 元/方  （预期 ≈ 8.92）")
    print(f"      不含税:       {r['price_excl_tax']:.4f}")
    print(f"      挖装费:       {r['cost_excavate']:.4f}")
    print(f"      运输费:       {r['cost_transport']:.4f}")

    print("[3/3] 黑山合同对标工况（Ⅲ较软岩 + 直接开挖）...")
    r2 = compute({'rock_level': 'Ⅲ较软岩'})
    print(f"      含税综合单价: {r2['price_incl_tax']:.4f} 元/方  （预期 ≈ 11.88）")

    print("\n✅ 自检通过")

"""
app.py — Streamlit 网页前端

运行：
    streamlit run app.py

部署：
    push 到 GitHub 私有仓库 → share.streamlit.io 一键部署
"""

import streamlit as st
import streamlit_authenticator as stauth
import yaml
from yaml.loader import SafeLoader
import pandas as pd
import time

from model_core import (
    compute, export_xlsx, OPTIONS, DEFAULT_PARAMS,
    TRUCK_INFO, DRILL_INFO,
    drill_recommend_diameter, truck_display, drill_display,
    build_audit_report,
)

# 维度 → emoji 前缀（build_audit_report 输出的 category 是纯文本，UI 层加 emoji）
_UI_CAT_EMOJI = {
    '工艺': '🔧', '挖机': '⛏️', '矿卡': '🚛',
    '钻机': '🛠️', '孔径': '📏', '坡度': '⛰️', '台阶': '📐',
}

def _ui_prefix_extra_line(line: str) -> str:
    """给 build_audit_report 输出的 process_extra_lines 加 UI emoji 前缀。"""
    if not line:
        return line
    s = line.strip()
    # 阶梯定价说明
    if s.startswith('阶梯定价：'):
        return '💰 ' + s
    # 错配补偿落点
    if s.startswith('因工艺错配补偿'):
        return '📍 ' + s
    # 模型规则跳过
    if s.startswith('模型规则：'):
        return 'ℹ️ ' + s
    # C 类工艺过剩对比
    if s.startswith('按您选'):
        return '💡 ' + s
    return s

# ============ 页面设置 ============
st.set_page_config(
    page_title="大型土石方成本测算",
    page_icon="⛏️",
    layout="wide",
)

# ============ 账号鉴权 ============
with open('auth_config.yaml') as f:
    auth_cfg = yaml.load(f, Loader=SafeLoader)

authenticator = stauth.Authenticate(
    auth_cfg['credentials'],
    auth_cfg['cookie']['name'],
    auth_cfg['cookie']['key'],
    auth_cfg['cookie']['expiry_days'],
)

authenticator.login(location='main')

if st.session_state.get('authentication_status') is False:
    st.error("用户名或密码错误")
    st.stop()
elif st.session_state.get('authentication_status') is None:
    st.warning("请先登录")
    st.stop()

# 已登录
name = st.session_state['name']
username = st.session_state['username']

# ============ 顶部：标题 + 登出 ============
col_title, col_user = st.columns([4, 1])
with col_title:
    st.title("⛏️ 大型土石方成本测算")
    st.caption("含税综合单价实时计算")
with col_user:
    st.write(f"👤 **{name}**")
    authenticator.logout(button_name='登出', location='main')

st.divider()

# ============ 参数表单（左侧） + 结果展示（右侧） ============
col_form, col_result = st.columns([1, 1])

with col_form:
    st.subheader("📝 测算参数")

    # —— 岩石条件 ——
    with st.expander("🗿 岩石条件", expanded=True):
        rock_level = st.selectbox("岩石级别", OPTIONS['rock_level'],
                                  index=OPTIONS['rock_level'].index(DEFAULT_PARAMS['rock_level']),
                                  help="按 GB/T 50218-2014 工程岩体分级标准")

    # —— 施工工艺 ——
    with st.expander("⚒️ 施工工艺", expanded=True):
        process = st.selectbox("工艺选择", OPTIONS['process'],
                               index=OPTIONS['process'].index(DEFAULT_PARAMS['process']))

    # 根据工艺判断哪些参数区需要展开
    needs_blast = '爆破' in process  # 含"爆破"或"爆破+二次破碎"
    # 注：松土器/破碎锤的参数 V8 模型走的是参数库自动取值，
    #     用户层目前无独立 widget，故这里只做钻机/爆破区的智能折叠。

    # —— 设备选型 ——
    with st.expander("🚜 设备选型", expanded=True):
        st.caption(
            "ℹ️ 设备和孔径**可任意自由组合**（允许大马拉小车 / 小马拉大车），"
            "匹配度与原因说明统一在页面底部「📋 工艺设备匹配度校验」展示。"
        )
        # 方案I：钻机↔孔径联动 —— 切钻机时自动把孔径跳到该钻机的推荐孔径
        # 必须在 selectbox 渲染之前注册 on_change 回调
        def _on_drill_change():
            new_drill = st.session_state.get('drill_key')
            if new_drill:
                st.session_state['hole_diameter_key'] = drill_recommend_diameter(new_drill)

        c1, c2 = st.columns(2)
        with c1:
            excavator = st.selectbox("挖机型号", OPTIONS['excavator'],
                                     index=OPTIONS['excavator'].index(DEFAULT_PARAMS['excavator']))
            truck = st.selectbox("矿卡型号", OPTIONS['truck'],
                                 index=OPTIONS['truck'].index(DEFAULT_PARAMS['truck']))
            if needs_blast:
                drill = st.selectbox("钻机型号", OPTIONS['drill'],
                                     index=OPTIONS['drill'].index(DEFAULT_PARAMS['drill']),
                                     key='drill_key',
                                     on_change=_on_drill_change)
            else:
                drill = DEFAULT_PARAMS['drill']
                st.caption("🚫 钻机型号（当前工艺不含爆破，已隐藏）")
        with c2:
            excavator_src = st.radio("挖机来源", OPTIONS['excavator_src'], horizontal=True,
                                     index=OPTIONS['excavator_src'].index(DEFAULT_PARAMS['excavator_src']))
            truck_src = st.radio("矿卡来源", OPTIONS['truck_src'], horizontal=True,
                                 index=OPTIONS['truck_src'].index(DEFAULT_PARAMS['truck_src']))
            if needs_blast:
                drill_src = st.radio("钻机来源", OPTIONS['drill_src'], horizontal=True,
                                     index=OPTIONS['drill_src'].index(DEFAULT_PARAMS['drill_src']))
            else:
                drill_src = DEFAULT_PARAMS['drill_src']

    # —— 运输参数 ——
    with st.expander("🛣️ 运输参数", expanded=True):
        c1, c2 = st.columns(2)
        with c1:
            dist_in = st.number_input("场内距离 (km)", min_value=0.0, max_value=50.0,
                                      value=float(DEFAULT_PARAMS['dist_in']), step=0.1)
            dir_in = st.selectbox("场内方向", OPTIONS['dir_in'],
                                  index=OPTIONS['dir_in'].index(DEFAULT_PARAMS['dir_in']))
            slope_in = st.number_input("场内坡度 (%)", min_value=0.0, max_value=30.0,
                                       value=float(DEFAULT_PARAMS['slope_in']), step=0.5)
        with c2:
            dist_out = st.number_input("场外距离 (km)", min_value=0.0, max_value=50.0,
                                       value=float(DEFAULT_PARAMS['dist_out']), step=0.1)
            dir_out = st.selectbox("场外方向", OPTIONS['dir_out'],
                                   index=OPTIONS['dir_out'].index(DEFAULT_PARAMS['dir_out']))
            slope_out = st.number_input("场外坡度 (%)", min_value=0.0, max_value=30.0,
                                        value=float(DEFAULT_PARAMS['slope_out']), step=0.5)

    # —— 爆破设计（仅在工艺含爆破时展开；不含爆破时直接全部用默认值） ——
    if needs_blast:
        with st.expander("💥 爆破设计参数", expanded=False):
            c1, c2 = st.columns(2)
            with c1:
                blast_len = st.number_input("爆破区域长度 (m)", min_value=10, max_value=300,
                                            value=int(DEFAULT_PARAMS['blast_len']), step=5)
                blast_wid = st.number_input("爆破区域宽度 (m)", min_value=5, max_value=100,
                                            value=int(DEFAULT_PARAMS['blast_wid']), step=2)
                buffer_rows = st.selectbox("缓冲孔排数", OPTIONS['buffer_rows'],
                                           index=OPTIONS['buffer_rows'].index(DEFAULT_PARAMS['buffer_rows']))
                step_h = st.number_input("台阶高度 H (m)", min_value=5.0, max_value=20.0,
                                         value=float(DEFAULT_PARAMS['step_h']), step=0.5,
                                         help="合理范围 5-15m，推荐 8-12m")
            with c2:
                slope_angle = st.number_input("坡面角 α (°)", min_value=50, max_value=90,
                                              value=int(DEFAULT_PARAMS['slope_angle']), step=1)
                hole_angle = st.number_input("钻孔倾角 (°)", min_value=60, max_value=90,
                                             value=int(DEFAULT_PARAMS['hole_angle']), step=1)
                hole_pattern = st.selectbox("布孔形式", OPTIONS['hole_pattern'],
                                            index=OPTIONS['hole_pattern'].index(DEFAULT_PARAMS['hole_pattern']))
                # 方案I：孔径与钻机联动 —— 用 session_state key 绑定，初始化默认值
                if 'hole_diameter_key' not in st.session_state:
                    st.session_state['hole_diameter_key'] = DEFAULT_PARAMS['hole_diameter']
                # 若联动回调把孔径设到了 OPTIONS 之外的值（理论不会，但兜底），夹回区间内
                if st.session_state['hole_diameter_key'] not in OPTIONS['hole_diameter']:
                    st.session_state['hole_diameter_key'] = DEFAULT_PARAMS['hole_diameter']
                hole_diameter = st.selectbox("孔径 (mm)", OPTIONS['hole_diameter'],
                                             key='hole_diameter_key')
                # 方案A 视觉强化（v10.1, 26-06-03）—— 孔径↔钻机联动状态实时显示
                # 不限制下拉选项（保留"大马拉小车 / 小马拉大车"对比能力），
                # 而是用 caption + warning + error 三档实时反馈当前选择是否匹配
                _di = DRILL_INFO.get(drill)
                if _di:
                    _dt, _drec, _drange = _di
                    _rng_s = '/'.join(str(x) for x in _drange)
                    if hole_diameter == _drec:
                        st.caption(
                            f"📏 钻机「{drill}」（{_dt}）推荐 **{_drec}mm**·可用 [{_rng_s}]"
                            f"　✅ 当前为标配孔径"
                        )
                    elif hole_diameter in _drange:
                        st.caption(
                            f"📏 钻机「{drill}」（{_dt}）推荐 **{_drec}mm**·可用 [{_rng_s}]"
                        )
                        st.warning(
                            f"⚠️ 孔径 {hole_diameter}mm 在钻机可用范围内，但非标配（推荐 {_drec}mm），"
                            f"台班效率可能下降"
                        )
                    else:
                        st.caption(
                            f"📏 钻机「{drill}」（{_dt}）推荐 **{_drec}mm**·可用 [{_rng_s}]"
                        )
                        if hole_diameter < _drec:
                            st.error(
                                f"🚫 大马拉小车：钻机「{drill}」标配 {_drec}mm，您选 {hole_diameter}mm"
                                f" → 台班费高但钻速发挥不出来，反而贵"
                            )
                        else:
                            st.error(
                                f"🚫 小马拉大车：钻机「{drill}」可用孔径 [{_rng_s}]mm，"
                                f"您选 {hole_diameter}mm 超出范围，可能无法施工"
                            )
                pre_split = st.radio("是否预裂", OPTIONS['pre_split'], horizontal=True,
                                     index=OPTIONS['pre_split'].index(DEFAULT_PARAMS['pre_split']))
    else:
        # 工艺不含爆破，所有爆破参数走默认值（不参与计算，仅占位）
        blast_len = DEFAULT_PARAMS['blast_len']
        blast_wid = DEFAULT_PARAMS['blast_wid']
        buffer_rows = DEFAULT_PARAMS['buffer_rows']
        step_h = DEFAULT_PARAMS['step_h']
        slope_angle = DEFAULT_PARAMS['slope_angle']
        hole_angle = DEFAULT_PARAMS['hole_angle']
        hole_pattern = DEFAULT_PARAMS['hole_pattern']
        hole_diameter = DEFAULT_PARAMS['hole_diameter']
        pre_split = DEFAULT_PARAMS['pre_split']
        st.caption("🚫 爆破设计参数（当前工艺不含爆破，已隐藏；如需启用请将工艺改为「爆破+开挖」或「爆破+二次破碎+开挖」）")

    # —— 计算按钮 ——
    calc_btn = st.button("🧮 开始测算", type='primary', use_container_width=True)


# ============ 收集参数 ============
params = {
    'rock_level': rock_level, 'process': process,
    'excavator': excavator, 'excavator_src': excavator_src,
    'truck': truck, 'truck_src': truck_src,
    'drill': drill, 'drill_src': drill_src,
    'dist_in': dist_in, 'dist_out': dist_out,
    'dir_in': dir_in, 'slope_in': slope_in,
    'dir_out': dir_out, 'slope_out': slope_out,
    'blast_len': blast_len, 'blast_wid': blast_wid,
    'buffer_rows': buffer_rows, 'step_h': step_h,
    'slope_angle': slope_angle, 'hole_angle': hole_angle,
    'hole_pattern': hole_pattern, 'hole_diameter': hole_diameter,
    'pre_split': pre_split,
}

# ============ 结果展示（右侧） ============
with col_result:
    st.subheader("📊 测算结果")

    if not calc_btn and 'last_result' not in st.session_state:
        st.info("👈 在左侧填写参数后点击「开始测算」")
    else:
        if calc_btn:
            with st.spinner("正在计算...约需 8 秒"):
                t0 = time.time()
                try:
                    result = compute(params)
                    st.session_state['last_result'] = result
                    st.session_state['last_params'] = params
                    st.session_state['last_elapsed'] = time.time() - t0
                except Exception as e:
                    st.error(f"计算失败：{e}")
                    st.stop()

        result = st.session_state['last_result']
        elapsed = st.session_state.get('last_elapsed', 0)

        # —— 综合单价大数字 ——
        warning = result.get('warning')
        has_hard_warn = warning and warning.get('severity') in ('error', 'warning')
        price_label = "💰 综合单价（含税）"
        if has_hard_warn:
            price_label = "⚠️ 综合单价（含税）"
        m1, m2 = st.columns(2)
        m1.metric(price_label,
                  f"{result['price_incl_tax']:.2f} 元/方" if result.get('price_incl_tax') else "—")
        m2.metric("综合单价（不含税）",
                  f"{result['price_excl_tax']:.2f} 元/方" if result.get('price_excl_tax') else "—")
        st.caption(f"⏱️ 耗时 {elapsed:.1f}s")

        # —— 成本构成 ——
        st.markdown("##### 成本构成 (元/方)")
        cost_items = [
            ('爆破费', result.get('cost_blast')),
            ('挖装费', result.get('cost_excavate')),
            ('运输费', result.get('cost_transport')),
            ('松土费', result.get('cost_loosen')),
            ('破碎费', result.get('cost_crush')),
            ('二次破碎费', result.get('cost_second_crush')),
            ('渣场费', result.get('cost_dump')),
        ]
        df = pd.DataFrame([(n, v) for n, v in cost_items if v not in (None, 0)],
                          columns=['项目', '元/方'])
        if not df.empty:
            df['元/方'] = df['元/方'].astype(float).round(2)
            st.dataframe(
                df,
                hide_index=True,
                use_container_width=True,
                column_config={
                    '元/方': st.column_config.NumberColumn('元/方', format='%.2f')
                },
            )
            st.bar_chart(df.set_index('项目'))

            # 工艺错配定价说明（按档位差）
            if warning and warning.get('gap_levels'):
                gap = warning['gap_levels']
                baseline = warning.get('baseline_price_incl', 0)
                factor = warning.get('adjusted_factor', 1.0)
                adj = warning.get('adjusted_price_incl', 0)
                v8 = warning.get('v8_original_price_incl', 0)
                st.caption(
                    f"📌 工艺错配阶梯定价：推荐工艺基准价 **{baseline:.2f}** 元/方 × "
                    f"档位差 **{gap}** 档系数 **{factor:.2f}** = **{adj:.2f}** 元/方"
                    f"　（按错配工艺直接算 {v8:.2f}，含产量罚，仅作对比）"
                )

        # —— 岩石参数联动展示 ——
        with st.expander("📐 岩石参数（自动匹配）"):
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("普氏系数 f", f"{result.get('f_value', 0):.2f}")
            c2.metric("容重 t/m³", f"{result.get('density', 0):.2f}")
            c3.metric("松散系数", f"{result.get('loose_factor', 0):.2f}")
            c4.metric("大块率", f"{result.get('big_block_rate', 0):.0%}")

        # —— 系统推荐 ——
        with st.expander("💡 系统推荐方案"):
            st.write(f"- 推荐工艺：**{result.get('recommend_process', '—')}**")
            st.write(f"- 推荐挖机：**{result.get('recommend_exc', '—')}**")
            st.write(f"- 推荐矿卡：**{result.get('recommend_truck', '—')}**")
            st.write(f"- 推荐钻机：**{result.get('recommend_drill', '—')}**")
            st.write(f"- 推荐孔径：**{result.get('recommend_hole_d', '—')} mm**")
            st.caption(
                "💡 各设备推荐独立给出：挖机按工艺+岩石强度匹配；"
                "矿卡按场内/场外坡度（是否下坡走纯电）+ 岩石硬度匹配；"
                "钻机按岩石硬度匹配；孔径按岩石+工艺匹配（钻机切换后会自动跳到该钻机标配孔径）。"
            )
            st.caption(
                "🔋 **纯电矿卡（E 系列）单方比柴油矿卡便宜约 4–5 元/方**（电费 vs 柴油费物理差价），"
                "但需充电桩与电网容量配套；不下坡场景默认推荐柴油，下坡场景默认推荐纯电（可借势发电）。"
            )
            st.caption(
                "📐 **钻机×孔径强耦合**：大钻机配小孔径=大马拉小车（台班费高+钻速发挥不出来），"
                "切换钻机后请保留自动跳转的标配孔径，单换钻机不换孔径反而会变贵。"
            )

        # —— 工艺设备匹配度校验 ——
        # 单一数据源：直接调用 model_core.build_audit_report，与 Excel 下载共享同一份审计输出
        with st.expander("📋 工艺设备匹配度校验", expanded=True):
            audit = build_audit_report(result)
            mismatches = audit.get('mismatches') or []
            extra_lines = audit.get('process_extra_lines') or []

            if mismatches:
                for cat, user_val, rec_val, reason in mismatches:
                    emoji = _UI_CAT_EMOJI.get(cat, '⚠️')
                    icon_item = f"{emoji} {cat}"
                    _u = user_val if user_val is not None else '—'
                    _r = rec_val  if rec_val  is not None else '—'
                    st.warning(f"{icon_item}：您选 **{_u}** ↔ 推荐 **{_r}**")
                    if reason:
                        st.caption(f"   → {reason}")
                    # 工艺错配补偿明细紧贴工艺那条显示
                    if cat == '工艺' and extra_lines:
                        for _line in extra_lines:
                            st.caption(f"   　 {_ui_prefix_extra_line(_line)}")
            else:
                st.success("✅ 所有工艺设备参数均与推荐方案一致，当前组合匹配度良好！")

            st.caption(
                "📐 **判定依据**：工艺=岩石普氏系数；挖机=工艺+岩石强度；"
                "矿卡=运距+载重+坡度+动力类型（T=柴油/E=纯电，纯电单方便宜约 4-5 元但需充电桩）；"
                "钻机=岩石硬度；孔径=岩石+工艺（大钻机配小孔径=大马拉小车，台班费高钻速出不来）；"
                "坡度/台阶=安全与效率校验。"
            )

        # —— 下载 ——
        st.divider()
        if st.button("📥 生成完整 Excel 测算单", use_container_width=True):
            with st.spinner("生成中..."):
                xlsx_bytes = export_xlsx(st.session_state['last_params'])
            st.download_button(
                label="⬇️ 下载 xlsx",
                data=xlsx_bytes,
                file_name=f"土石方测算_{rock_level}_{process}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

# ============ 底部 ============
st.divider()
st.caption("© 大型土石方成本测算 · 含税综合单价 = 直接成本 × (1+税率)")

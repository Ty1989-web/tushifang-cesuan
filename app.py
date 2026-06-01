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

from model_core import compute, export_xlsx, OPTIONS, DEFAULT_PARAMS

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
    st.caption("基于 V8 模型 · 含税综合单价实时计算")
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
        c1, c2 = st.columns(2)
        with c1:
            excavator = st.selectbox("挖机型号", OPTIONS['excavator'],
                                     index=OPTIONS['excavator'].index(DEFAULT_PARAMS['excavator']))
            truck = st.selectbox("矿卡型号", OPTIONS['truck'],
                                 index=OPTIONS['truck'].index(DEFAULT_PARAMS['truck']))
            if needs_blast:
                drill = st.selectbox("钻机型号", OPTIONS['drill'],
                                     index=OPTIONS['drill'].index(DEFAULT_PARAMS['drill']))
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
                hole_diameter = st.selectbox("孔径 (mm)", OPTIONS['hole_diameter'],
                                             index=OPTIONS['hole_diameter'].index(DEFAULT_PARAMS['hole_diameter']))
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
                    f"　（V8 按错配工艺算 {v8:.2f}，含 V8 自带产量罚，仅作对比）"
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
            st.write(f"- 推荐钻机：**{result.get('recommend_drill', '—')}**")

        # —— 校验警告 ——
        warns = [v for k, v in result.items() if k.startswith('warn_') and v and '⚠' in str(v)]

        # 1) 工艺错配警告（最显眼）
        if warning:
            sev = warning.get('severity')
            ptype = warning.get('process_user', '—')
            rec = warning.get('process_recommend', '—')
            msg = warning.get('message', '')
            gap = warning.get('gap_levels', 0)
            factor = warning.get('adjusted_factor')
            baseline = warning.get('baseline_price_incl')
            adj = warning.get('adjusted_price_incl')
            v8 = warning.get('v8_original_price_incl')

            # 拼定价说明（gap > 0 才有）
            price_detail = ""
            if gap and factor:
                price_detail = (
                    f"\n- 阶梯定价：推荐工艺基准 **{baseline:.2f}** × 档位差 **{gap}** 档系数 **{factor:.2f}** = **{adj:.2f}** 元/方"
                    f"\n- V8 按错配工艺算（含自带产量罚）：{v8:.2f} 元/方"
                )

            if sev == 'error':
                st.error(
                    f"🚨 **工艺错配严重警告**：当前岩石推荐 **{rec}**，您选择了 **{ptype}**（偏软 {gap} 档）。\n\n"
                    f"- {msg}"
                    f"{price_detail}"
                    f"\n- ⚠️ **强烈建议改用「{rec}」**，否则成本将大幅上升。"
                )
            elif sev == 'warning':
                st.warning(
                    f"⚠️ **工艺错配警告**：当前岩石推荐 **{rec}**，您选择了 **{ptype}**（偏软 {gap} 档）。\n\n"
                    f"- {msg}"
                    f"{price_detail}"
                    f"\n- 建议考虑改用「{rec}」。"
                )
            elif sev == 'info':
                if gap and gap > 0:
                    # 轻微偏软（如较软岩用直接开挖）
                    st.info(
                        f"💡 **工艺偏软提示**：当前岩石推荐 **{rec}**，您选择了 **{ptype}**（偏软 {gap} 档）。"
                        f"{('（' + msg + '）') if msg else ''}"
                        f"{price_detail}"
                    )
                else:
                    # C 类：工艺过剩/杀鸡牛刀（方案G 文案修复）
                    excess = warning.get('excess_levels', 0)
                    v8_actual = warning.get('v8_actual_price_incl')
                    rec_base = warning.get('recommend_baseline_incl')
                    overuse_detail = ""
                    if v8_actual is not None and rec_base is not None:
                        diff = v8_actual - rec_base
                        sign = '+' if diff >= 0 else ''
                        overuse_detail = (
                            f"\n- V8 按您选的「**{ptype}**」算实际成本：**{v8_actual:.2f}** 元/方（含税）"
                            f"\n- 推荐工艺「**{rec}**」默认参数下成本约 **{rec_base:.2f}** 元/方"
                            f"\n- 差异约 **{sign}{diff:.2f}** 元/方，源自 V8 按不同工艺算的物料参数（松散系数、大块率、运输方量等）"
                            f"\n- 校正层未额外加罚，**显示的就是 V8 按您选工艺算的真实成本**"
                        )
                    st.info(
                        f"💡 **工艺过剩提示**：当前岩石推荐 **{rec}**，您选择了 **{ptype}**"
                        f"（过剩 {excess} 档，杀鸡用牛刀）。"
                        f"{('（' + msg + '）') if msg else ''}"
                        f"{overuse_detail}"
                    )

            # 方案G 新增：B 类错配补偿落点说明（让用户看清楚罚加在了哪个成本项）
            penalty_label = warning.get('penalty_field_label')
            penalty_extra = warning.get('penalty_extra')
            note = warning.get('note')
            if penalty_label and penalty_extra is not None and penalty_extra > 0:
                st.caption(
                    f"📍 错配补偿落点：差价 **{penalty_extra:.2f}** 元/方（不含税）"
                    f"已加到「**{penalty_label}**」上，其他成本项保持 V8 原算真实值不变。"
                )

            # 方案G 新增：V8 模型跳过项解释（如极软岩+爆破工艺时爆破费=0）
            skip_reason = warning.get('v8_skip_reason')
            if skip_reason:
                st.caption(f"ℹ️ {skip_reason}")

        # 2) 其他校验提示（原 warn_* 字段）
        if warns:
            st.warning("⚠️ 校验提示：\n\n" + "\n\n".join(f"- {w}" for w in warns))

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
st.caption("© 大型土石方成本测算 · V8 模型 · 含税综合单价 = 直接成本 × (1+税率)")

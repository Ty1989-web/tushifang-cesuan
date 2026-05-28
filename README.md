# 大型土石方成本测算 · 网页版

基于 V8 Excel 模型的网页化部署，30-40 人小团队内部使用。

## 📁 目录结构

```
webapp/
├── model_core.py        ← 计算核心（封装 build_v8.py + libreoffice 重算）
├── app.py               ← Streamlit 前端
├── auth_config.yaml     ← 账号配置（4 个默认账号）
├── requirements.txt     ← Python 依赖
├── packages.txt         ← 系统依赖（仅 Streamlit Cloud 使用）
├── .gitignore
└── README.md            ← 本文件

../build_v8.py           ← V8 模型（被 model_core 引用，勿删）
```

## 🚀 本地试跑（5 分钟）

```bash
# 1. 安装依赖
pip install -r requirements.txt

# 2. 系统装 libreoffice
#    Ubuntu/Debian:  sudo apt install libreoffice
#    macOS:          brew install --cask libreoffice

# 3. 启动
cd webapp
streamlit run app.py

# 4. 浏览器自动打开 http://localhost:8501
#    用 admin / Admin#2026 登录
```

## 🌐 部署到 Streamlit Community Cloud（永久免费）

### 准备
- 一个 GitHub 账号（免费的就行）
- 一个 Streamlit 账号（用 GitHub 一键登录）

### 步骤

**1. 改安全配置**（重要！）

编辑 `auth_config.yaml`，修改两处：
- `cookie.key`：改成一段随机长字符串（防 cookie 伪造）
- `credentials.usernames.*.password`：每个账号都改成新密码的 hash

生成新密码 hash 的命令：
```bash
python3 -c "import bcrypt; print(bcrypt.hashpw(b'新密码', bcrypt.gensalt()).decode())"
```

**2. 推送到 GitHub**

```bash
cd 代码版本                    # 注意：要把 build_v8.py 一起推
git init
git add .
git commit -m "初版：大型土石方成本测算网页"
# 在 GitHub 网页创建一个【私有仓库】heishan-calc
git remote add origin git@github.com:你的用户名/heishan-calc.git
git push -u origin main
```

**3. 部署**

1. 打开 https://share.streamlit.io
2. 用 GitHub 登录
3. 点 "New app"
4. 配置：
   - Repository: `你的用户名/heishan-calc`
   - Branch: `main`
   - Main file path: `webapp/app.py`
5. 点 "Deploy"

约 3-5 分钟后，拿到 `https://heishan-calc.streamlit.app` 之类的链接。
把链接和账号密码发给 30-40 个同事即可使用。

## 👥 账号管理

### 添加新用户
编辑 `auth_config.yaml`，在 `credentials.usernames` 下增加：

```yaml
    zaojia3:
      name: 造价员三
      password: <bcrypt hash>
      email: zaojia3@example.com
      first_name: ''
      last_name: ''
```

`git push` 后 Streamlit Cloud 会自动重新部署，1-2 分钟生效。

### 默认账号清单

| 用户名 | 默认密码 | 用途 |
|---|---|---|
| admin | Admin#2026 | 管理员 |
| zaojia1 | Zj1#2026 | 造价员 |
| zaojia2 | Zj2#2026 | 造价员 |
| viewer | View#2026 | 只读查看 |

⚠️ **生产部署前务必全部改密码！**

## 🔧 改模型参数

不需要改任何 Excel，直接改代码：

| 改什么 | 改哪 |
|---|---|
| 主表公式 / 单价 / 系数 | `代码版本/build_v8.py` 里搜对应 cell（如 `('B65', ...)`） |
| 参数库的常量 | `build_v8.py` 中搜对应 sheet name 段落 |
| 网页表单字段 | `webapp/app.py` |
| 下拉选项（如新增挖机型号） | `webapp/model_core.py` 中的 `OPTIONS` 字典 |

改完后：
- 本地：删 `webapp/template_v8.xlsx` 后重启 streamlit
- 线上：`git push` 自动重部署

## ⚡ 性能说明

- 每次「开始测算」约 **6-10 秒**（libreoffice 重算 + IO）
- Streamlit Cloud 免费版：1 GB 内存、7 天无访问休眠（冷启动 30 秒）
- 30-40 人偶尔使用完全够，并发上百再考虑升级

## 🆘 常见问题

**Q：登录后页面一直转圈？**
A：第一次冷启动需要 30 秒，是 Streamlit Cloud 休眠唤醒机制，正常。

**Q：计算结果和 Excel 不一致？**
A：检查 `build_v8.py` 是不是最新版；删除 `webapp/template_v8.xlsx` 强制重建模板。

**Q：怎么关停 / 删除应用？**
A：share.streamlit.io 后台对应 app 旁边有「⋮」菜单 → Delete。

**Q：免费版用户数有限制吗？**
A：没有用户数限制，但**单 app 同时在线一般几十人没问题**，超过会排队。

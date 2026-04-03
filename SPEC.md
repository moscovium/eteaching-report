# E听说 成效报告系统 — Streamlit Cloud 部署说明

## 📋 仓库文件清单

```
eteaching-report/
├── report_app.py          # Streamlit 主程序（唯一入口）
├── requirements.txt       # Python 依赖版本锁定
├── SPEC.md                # 本文件
└── README.md              # 用户使用说明
```

## 🚀 部署到 Streamlit Cloud（免费）

### 步骤一：创建 GitHub 仓库

1. 打开 [https://github.com/new](https://github.com/new)
2. **Repository name** 填：`eteaching-report`
3. 选择 **Public**（Streamlit Cloud 只能部署公开仓库）
4. 点击 **Create repository**（不要勾选任何初始化选项）

### 步骤二：推送代码到新仓库

在终端运行（把 `USERNAME` 换成你的 GitHub 用户名）：

```bash
cd /Users/x/Downloads/Project

# 添加新仓库为另一个 remote（避免和现有的 origin 冲突）
git remote add gh https://github.com/USERNAME/eteaching-report.git

# 推送到新仓库的 main 分支
git subtree split --prefix=. --branch gh-pages 2>/dev/null || true
git push gh gh-pages:main
```

> ⚠️ **如果以上命令报错**（subtree 较复杂），可以改用更简单的方式：
> ```bash
> # 在 GitHub 页面手动上传文件
> # 或用 GitHub CLI：
> gh repo create eteaching-report --public --source=. --push
> ```

### 步骤三：在 Streamlit Cloud 部署

1. 打开 [https://share.streamlit.io](https://share.streamlit.io)
2. 用 GitHub 账号登录（首次需要授权）
3. 点击 **New app**
4. 填写：
   - **Repository**：`你的GitHub用户名/eteaching-report`
   - **Branch**：`main`
   - **Main file path**：`report_app.py`
5. 点击 **Deploy!**

### 步骤四：获取访问链接

部署成功后，Streamlit 会给你一个链接，格式如：
`https://eteaching-report.streamlit.app`

把这个链接分享给同事即可，**完全免费，无需对方配置任何环境**。

---

## 🔧 注意事项

### 1. Streamlit Cloud 限制
- 应用在 **休眠机制**：超过7天无访问会休眠（唤醒需要约30秒）
- 每个账号有 **免费额度**：每月 1000 小时活跃时间
- 文件上传限制：单个文件不超过 200MB

### 2. 数据隐私
- 数据处理在 Streamlit Cloud 服务器上进行
- 如果数据敏感，建议慎用（考虑本地部署方案）

### 3. Streamlit 版本
- `report_app.py` 使用 Streamlit 1.x API
- 所有依赖版本已在 `requirements.txt` 中锁定

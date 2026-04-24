# GitHub 手动上传指南

## 当前状态

Git 本地提交已完成，但无法连接到 GitHub 远程服务器进行推送。

## 本地提交记录

```
b4376f7 v3.1: Sheet页选择、过滤配置、异常分析显示修复、数据校验、斜杠分隔符、弹窗使用说明
```

## 更新内容

| 文件 | 变更 |
|------|------|
| `traceability_v3_1.py` | 主程序更新（Sheet选择、过滤、弹窗等） |
| `requirements.txt` | 添加 pyinstaller 依赖 |
| `使用说明文档_v3.1.md` | 新增 v3.1 使用说明 |

## 手动上传到 GitHub 步骤

### 方法 1：GitHub Web 界面直接上传

1. 打开浏览器，访问 https://github.com/yysbnzy/traceability-tool
2. 点击 **Add file** → **Upload files**
3. 选择以下文件上传：
   - `traceability_v3_1.py`
   - `requirements.txt`
   - `使用说明文档_v3.1.md`
4. 填写提交信息：
   ```
   v3.1: Sheet页选择、过滤配置、异常分析显示修复、数据校验、斜杠分隔符、弹窗使用说明
   ```
5. 点击 **Commit changes**

### 方法 2：GitHub Desktop（推荐）

1. 下载安装 [GitHub Desktop](https://desktop.github.com/)
2. 登录 GitHub 账号
3. 添加本地仓库：`File` → `Add local repository`
4. 选择 `C:\Users\Administrator\.openclaw\workspace\projects\traceability-tool`
5. 点击 **Publish branch** 或 **Push origin**

### 方法 3：命令行重试

等网络恢复后，在 PowerShell 中执行：

```powershell
cd C:\Users\Administrator\.openclaw\workspace\projects\traceability-tool
git push origin master
```

如果提示输入用户名密码：
- 用户名：你的 GitHub 用户名
- 密码：GitHub Personal Access Token（不是登录密码）

## 生成新的 GitHub Token

1. 打开 https://github.com/settings/tokens
2. 点击 **Generate new token** → **Generate new token (classic)**
3. 填写 Note：`traceability-tool-push`
4. 勾选权限：**repo**（完整仓库访问）
5. 点击 **Generate token**
6. 复制生成的 token（只显示一次）
7. 更新 git remote URL：
   ```powershell
   cd C:\Users\Administrator\.openclaw\workspace\projects\traceability-tool
   git remote set-url origin https://你的token@github.com/yysbnzy/traceability-tool.git
   git push origin master
   ```

## 文件清单（v3.1 更新）

已打包到 `dist/` 目录：
- `溯源工具_v3_1.exe` - 主程序
- `溯源工具_v3.1_使用说明.docx` - 使用说明文档

---
*生成时间：2026-04-23*
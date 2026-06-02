# GitHub 推送到 Netlify 发布指南

本文档适用于新版“学生时间规划表图片生成器”：

- 前端网页在 `timeline_web/`
- Netlify Function 在 `timeline_web/netlify/functions/`
- 需要读取飞书在线表格
- 需要配置飞书 `App ID` 和 `App Secret`

推荐做法：把 `timeline_web/` 作为一个独立 GitHub 仓库的根目录。

## 一、准备 GitHub 仓库

### 1. 新建仓库

打开 GitHub：

```text
https://github.com/new
```

你的仓库：

```text
https://github.com/gaohjwork-hub/student-timeline-web
```

建议选择：

```text
Private
```

不要勾选：

```text
Add a README file
Add .gitignore
Choose a license
```

因为本地项目里已经有文件。

仓库地址：

```text
https://github.com/gaohjwork-hub/student-timeline-web.git
```

## 二、把本地项目推送到 GitHub

打开终端，进入项目目录：

```bash
cd /Users/gaohuijun/Documents/杭星/timeline_web
```

初始化 Git：

```bash
git init
```

添加文件：

```bash
git add .
```

提交：

```bash
git commit -m "Initial Netlify Feishu app"
```

设置主分支：

```bash
git branch -M main
```

绑定 GitHub 仓库：

```bash
git remote add origin https://github.com/gaohjwork-hub/student-timeline-web.git
```

如果本地已经存在 `origin`，改用：

```bash
git remote set-url origin https://github.com/gaohjwork-hub/student-timeline-web.git
```

推送：

```bash
git push -u origin main
```

如果 GitHub 要求登录：

- 可以用 GitHub Desktop
- 或使用 GitHub Personal Access Token
- 或按终端提示完成浏览器授权

## 三、在 Netlify 连接 GitHub 仓库

打开 Netlify：

```text
https://app.netlify.com/
```

进入：

```text
Add new site -> Import an existing project
```

选择：

```text
GitHub
```

授权 Netlify 访问你的 GitHub 后，选择仓库：

```text
student-timeline-web
```

## 四、Netlify 构建设置

如果你是按推荐方式，把 `timeline_web/` 作为仓库根目录，Netlify 设置如下：

```text
Base directory: 留空
Build command: 留空
Publish directory: .
Functions directory: netlify/functions
```

项目里已经有：

```text
netlify.toml
```

Netlify 通常会自动读取：

```toml
[build]
  publish = "."
  functions = "netlify/functions"
```

如果 Netlify 页面上没有 Functions directory 输入框，也没关系，`netlify.toml` 会处理。

点击：

```text
Deploy site
```

等待首次部署完成。

## 五、配置 Netlify 环境变量

部署完成后进入站点后台：

```text
Site configuration -> Environment variables
```

新增两个变量：

```text
FEISHU_APP_ID=你的飞书 App ID
FEISHU_APP_SECRET=你的飞书 App Secret
```

保存。

然后重新部署：

```text
Deploys -> Trigger deploy -> Deploy site
```

说明：

- 如果没有重新部署，Netlify Function 可能读取不到刚刚新增的环境变量。
- 只有 GitHub 连接的站点通常才会看到 `Trigger deploy`。
- 手动拖拽 zip 的站点通常没有这个按钮。

## 六、确认 Netlify Function 是否存在

部署后，复制你的 Netlify 域名，例如：

```text
https://student-timeline-web.netlify.app/
```

在浏览器打开：

```text
https://student-timeline-web.netlify.app/.netlify/functions/feishu-sheets
```

正常情况下会看到类似：

```json
{"error":"缺少飞书表格 token 或链接。"}
```

这说明 Function 已经部署成功。

如果看到 404：

- 检查仓库里是否有 `netlify/functions/feishu-sheets.js`
- 检查 `netlify.toml` 是否存在
- 检查 Netlify 是否连接的是正确仓库
- 检查 Publish directory 是否是 `.`

## 七、配置飞书自建应用

进入飞书开放平台：

```text
https://open.feishu.cn/app
```

进入你的企业自建应用。

### 1. 添加网页应用能力

路径：

```text
应用能力 -> 添加应用能力 -> 网页应用
```

主页地址填写 Netlify 地址：

```text
https://student-timeline-web.netlify.app/
```

### 2. 开通权限

路径：

```text
权限管理
```

搜索并开通和以下含义对应的权限：

```text
读取电子表格
查看电子表格
查看云空间中文件
```

权限名称可能因飞书后台版本不同略有差异。核心原则是：应用必须能读取云文档中的电子表格。

### 3. 发布应用版本

路径：

```text
应用发布 -> 版本管理与发布 -> 创建版本
```

版本说明建议：

```text
支持读取飞书在线电子表格子表，并生成可编辑时间规划图。
```

提交审核。

## 八、飞书在线表格授权

如果读取子表失败，除了应用权限，还要检查目标在线表格本身的访问权限。

建议测试时先做：

- 确认你本人能打开该在线表格
- 确认表格属于同一个飞书企业
- 确认应用已发布并审核通过
- 如后台支持，将应用加入文档可访问范围

## 九、上线测试

打开 Netlify 网页或飞书工作台里的应用，测试：

1. 粘贴飞书在线电子表格链接
2. 点击“读取子表”
3. 选择子表
4. 点击“生成表格”
5. 检查标题、阶段、年级、月份
6. 检查底表颜色是否按“暑期、上学期、寒假、下学期”循环
7. 编辑色块文字
8. 拖动或拉伸色块
9. 下载 PNG
10. 下载 PDF

## 十、以后如何更新应用

每次我帮你改完 `timeline_web/` 里的文件后，你在终端执行：

```bash
cd /Users/gaohuijun/Documents/杭星/timeline_web
git add .
git commit -m "Update timeline app"
git push
```

Netlify 会自动检测 GitHub 更新并重新部署。

你不需要再手动拖 zip。

## 十一、如果你把整个“杭星”目录推到 GitHub

不推荐，但也可以。

如果仓库根目录是：

```text
/Users/gaohuijun/Documents/杭星
```

那 Netlify 构建设置要改为：

```text
Base directory: timeline_web
Build command: 留空
Publish directory: .
Functions directory: netlify/functions
```

但推荐保持简单：只把 `timeline_web/` 单独作为 GitHub 仓库。

## 十二、常见问题

### 1. Netlify 页面能打开，但读取飞书表格失败

检查：

- `FEISHU_APP_ID` 是否正确
- `FEISHU_APP_SECRET` 是否正确
- 环境变量保存后是否重新部署
- 飞书应用是否开通电子表格读取权限
- 飞书应用是否发布并审核通过
- 目标表格是否允许读取

### 2. Function 地址 404

检查：

- `netlify/functions/feishu-sheets.js` 是否推送到 GitHub
- `netlify.toml` 是否推送到 GitHub
- Netlify 是否连接了正确仓库
- Publish directory 是否为 `.`

### 3. `Trigger deploy` 找不到

如果是 GitHub 连接站点，通常在：

```text
Deploys -> Trigger deploy
```

如果还是找不到，可以直接修改一个文件后重新 `git push`，Netlify 会自动部署。

### 4. 本地 Excel 上传可用，但飞书在线表格不可用

这说明前端页面正常，问题集中在：

- Netlify Function
- Netlify 环境变量
- 飞书开放平台权限
- 在线表格访问权限

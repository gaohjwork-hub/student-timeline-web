# 学生时间规划表图片生成器

入口文件：

`index.html`

功能：

- 上传结构一致的学生时间规划表 Excel
- 粘贴飞书在线电子表格链接并读取子表
- 生成星屿风格时间规划表预览
- 按实际表格列数生成，不保留尾部空列
- 按内容自适应行高和列宽，正文字号保持统一
- 底表按暑期、上学期、寒假、下学期循环分色
- 点击任务色块直接编辑文字
- 新增、删除任务色块
- 用色块右上角圆点拖动位置
- 用色块四边短条拉伸到更多月份或任务列
- 下载 PNG 或 PDF

发布方式：

把整个 `timeline_web/` 目录上传到任意静态网页托管服务即可。

飞书自建应用发布步骤见：

`FEISHU_SETUP.md`

飞书在线表格读取 + Netlify Function 版本见：

`FEISHU_ONLINE_SHEETS_NETLIFY_GUIDE.md`

GitHub 推送到 Netlify 发布流程见：

`GITHUB_NETLIFY_DEPLOY_GUIDE.md`

外部依赖：

- SheetJS：读取 Excel
- jsPDF：导出 PDF

服务端函数：

- `netlify/functions/feishu-sheets.js`：读取飞书在线电子表格

说明：

导出不依赖网页截图，而是用预览中的同一份数据重新绘制 Canvas，因此输出稳定性更高。

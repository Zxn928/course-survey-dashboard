# 课程问卷数据看板

这是一个公开匿名版的课程问卷数据看板，用于展示线下课程反馈分析结果。

## 在线预览

[打开数据看板](https://zxn928.github.io/day2-course-survey-dashboard/)

## 看板内容

- 核心指标：有效答卷数、整体满意度、非常满意占比、可行动建议数
- 评分分析：教师评价与整体满意度分布
- 主题归纳：课程收获主题、建议主题与优化方向
- 原始反馈：逐条浏览反馈内容，并支持筛选有建议的答卷
- 移动端适配：可在手机浏览器中直接访问

## 数据隐私

发布版本已做匿名化处理：

- 不包含原始 CSV 中的姓名字段
- 不包含提交人字段
- 保留问卷编号、日期、评分、课程收获与意见建议，便于复盘分析

如果后续需要公开传播，建议继续使用匿名版本，避免上传包含个人身份信息的原始问卷文件。

## 项目文件

```text
.
├── index.html   # 自包含 HTML 数据看板
├── README.md    # 项目说明
└── .nojekyll    # GitHub Pages 静态站点配置
```

## 更新方式

如需用新的 CSV 或 Excel 重新生成看板，可以使用本地 Codex skill：

```bash
python3 ~/.codex/skills/survey-data-dashboard/scripts/build_survey_dashboard.py INPUT_FILE --output index.html --title "课程问卷数据看板"
```

生成新的 `index.html` 后，重新提交到 GitHub Pages 仓库即可更新线上页面。

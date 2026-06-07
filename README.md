# 课程问卷数据看板

一个用于课程课后问卷标准化分析的 Codex Skill。安装后，Agent 可以根据 Excel、CSV 或 TSV 问卷文件，自动完成数据解析、满意度统计、开放反馈主题归纳，并生成可外发的 HTML 可视化看板。

## 在线示例

[查看示例看板](https://zxn928.github.io/course-survey-dashboard/)

## 能做什么

- 解析课程问卷 CSV / TSV / XLSX 文件
- 自动识别提交时间、评分/满意度、课程收获、意见建议等字段
- 生成核心指标、评分分布、反馈主题、建议主题和原始反馈表
- 输出自包含 HTML 文件，适合发给老师、同学或项目成员查看
- 支持生成 zip 外发包，避免部分 IM/邮箱拦截 HTML 附件
- 提醒并处理 `file:///...` 本地路径无法外发的问题

## 安装

将本仓库克隆到 Codex skills 目录：

```bash
git clone https://github.com/Zxn928/course-survey-dashboard.git ~/.codex/skills/course-survey-dashboard
```

安装后，重新打开 Codex 会话或刷新技能列表，即可使用：

```text
使用 $course-survey-dashboard，帮我把这个课后问卷 Excel 生成 HTML 数据看板。
```

## 仓库结构

```text
.
├── SKILL.md                         # Skill 主说明文件
├── scripts/
│   └── build_survey_dashboard.py     # 通用 HTML 看板生成脚本
├── docs/
│   ├── index.html                    # GitHub Pages 示例看板
│   └── .nojekyll
└── README.md
```

## 使用脚本

也可以直接运行脚本生成看板：

```bash
python3 scripts/build_survey_dashboard.py INPUT_FILE --output OUTPUT_HTML --title "课程问卷数据看板" --zip
```

支持格式：

- `.csv`
- `.tsv`
- `.xlsx`，需要 Python 环境中安装 `openpyxl`

## 数据隐私

示例看板为公开匿名版：

- 不包含原始 CSV 中的姓名字段
- 不包含提交人字段
- 保留问卷编号、日期、评分、课程收获与意见建议，用于复盘分析

公开发布课程问卷时，建议默认使用匿名版本，避免上传包含个人身份信息的原始文件。

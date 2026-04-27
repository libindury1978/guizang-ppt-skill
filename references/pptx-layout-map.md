# PPTX Layout Map

将现有 HTML layout 映射到 `scripts/build_pptx.py` 的 `slide.type`，用于生成可编辑 `.pptx`。

| HTML Layout | 推荐 `slide.type` | 说明 |
|---|---|---|
| 1. 开场封面 | `cover` | 大标题 + 副标题 |
| 2. 章节幕封 | `section` | 章节切分页 |
| 3. 数据大字报 | `stats` | 3 个关键数字 |
| 4. 左文右图 | `content_image` | 左文本，右图片 |
| 5. 图片网格 | `image_grid` | 2-4 张图拼版 |
| 6. 两列流水线 | `pipeline` | 3-6 步流程 |
| 7. 悬念问题页 | `quote` | 重点问题/金句 |
| 8. 大引用页 | `quote` | 大号引述 |
| 9. 并列对比 | `comparison` | before vs after |
| 10. 图文混排 | `content_image` | 主图 + 文本要点 |

## JSON 输入示例

```json
{
  "title": "一种新的工作方式",
  "subtitle": "被 AI 重塑的团队协作",
  "author": "Benny",
  "theme": "ink-classic",
  "slides": [
    {
      "type": "cover",
      "title": "一种新的工作方式",
      "subtitle": "被 AI 重塑的团队协作"
    },
    {
      "type": "section",
      "title": "01 问题背景",
      "subtitle": "为什么现在必须改变"
    },
    {
      "type": "content_image",
      "title": "组织正在被折叠",
      "bullets": ["角色边界变薄", "决策回路变短", "执行速度提升"],
      "image": "images/03-team.jpg"
    }
  ]
}
```

## 字段约定

- `type`: 必填。支持 `cover` / `section` / `content_image` / `comparison` / `quote` / `stats` / `pipeline` / `image_grid`。
- `title`: 建议对应 L1。
- `subtitle`、`bullets`、`quote`、`stats`、`steps`: 按页面类型填。
- `image` / `images`: 图片相对路径，建议和 `deck.json` 同级管理。

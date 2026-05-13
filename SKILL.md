---
name: amazon-voc-tagger
description: 对亚马逊（或其他电商平台）产品评论进行 VOC（Voice of Customer）系统性打标签分析，生成三级标签体系、带标签的 Excel 和可交互 HTML 分析报告。使用场景：(1) 需要分析一批产品评论并结构化洞察，(2) 需要建立三级标签分类体系并逐条打标，(3) 需要生成可视化的 VOC 分析报告（评分趋势/问题地图/优先级矩阵），(4) 需要对比不同市场（US/JP等）的用户反馈差异。触发词：亚马逊评论分析、VOC分析、用户评价打标签、review分析、差评分析、用户之声、评论分类。
---

# Amazon VOC Tagger

对产品评论进行三级标签分类分析，输出带标签 Excel + 交互式 HTML 报告。

## 核心技术方案

**打标方式：CRS claude-sonnet-4-6 + tool call + enum 约束**

这是经过验证的唯一可靠方案。其他方案的失败原因：
- `response_format: json_object` → CRS 代理不透传，模型无视
- 纯 system prompt 要求输出 JSON → 模型返回 markdown 分析文章
- tool call 不加 enum → 模型自创标签体系，不使用我们的分类

**正确方案：** 在 tool 定义的 `l1/l2/l3` 字段加 `enum` 数组，枚举所有合法值，强制模型只能从中选择。

```python
'l1': {'type': 'string', 'enum': L1_ENUM},   # 所有一级标签列表
'l2': {'type': 'string', 'enum': ALL_L2},    # 所有二级标签列表（flat）
'l3': {'type': 'string', 'enum': ALL_L3},    # 所有三级标签列表（flat）
```

## 工作流程

### 第一步：建立标签体系

1. 读取 Excel，采样 50-100 条评论（含好评差评各半）
2. 与用户协作归纳一级标签（产品特性维度：稳定性/性能/售后等）
3. 细化二级标签（问题子类）
4. 定义三级标签（具体可执行的问题点，能直接对应改进项）
5. **输出标签体系后暂停，等待用户确认**，可能需要多轮调整

**标签体系原则（MECE）：**
- 一级：5-15 个宏观维度，覆盖正向和负向反馈
- 二级：每个一级下 2-6 个子类
- 三级：能直接对应产品可执行改进点，粒度够细
- 正向反馈单独设为「0.正向反馈」一级，下设好评亮点子类

### 第二步：打标签

运行 `scripts/tag_reviews.py`，需传入：
- 输入 Excel 路径
- 标签体系（编辑脚本内的 `TAXONOMY` 和 enum 变量）
- CRS API 配置

**多标签处理：** 一条评论多个问题 → 展开为多行，每行一个三级标签组合

**Checkpoint：** 每批完成后保存到 `tag_checkpoint.json`，网络中断可断点续跑

### 第三步：数据分析与报告

运行 `scripts/extract_data.py` 提取统计数据，再生成 HTML 报告。

报告包含（参考 `references/report-sections.md`）：
- 总体概况 KPI + 星级分布
- 月度评分变化趋势
- 差评一级分类分布（按可改善性着色）
- 三级标签 Top20 明细表
- 季度差评分布趋势
- 优先级矩阵（频次 × 可改善性四象限）
- 行动建议（短/中/长期）
- 完整差评分类明细（可展开透视表）

## 配置要求

```python
CRS_BASE = 'https://your-openai-compatible-endpoint/api'  # 任意 OpenAI-compatible 接口
CRS_KEY  = 'cr_...'          # 从 ~/.openclaw/secrets/openclaw.env 读取
PROXY    = 'http://127.0.0.1:6478'
MODEL    = 'claude-sonnet-4-6'
BATCH    = 10                 # 每批条数，建议 8-12
```

## 输入格式要求

Excel 至少包含以下字段：
- `id`：评论唯一标识
- `star`：评分（1-5）
- `market`：市场（US/JP/UK等）
- `title_zh` / `text_zh`：中文标题和正文（已翻译）
- `date` 或 `created_at`：评论时间（用于趋势分析）

## 输出

1. `{product}_tagged.xlsx`：原始字段 + `l1/l2/l3` 三列标签，多标签展开多行
2. `{product}_VOC_Report.html`：完整交互式分析报告，单文件可直接浏览器打开

## 常见问题

**Q: 批次失败返回空结果怎么办？**
脚本会自动重试3次，checkpoint 保存已完成批次，直接重跑脚本会跳过已完成的。

**Q: 标签有遗漏（某些评论内容体系覆盖不到）怎么办？**
打标完成后检查空标签行，单独补打时用更宽松的 prompt，或向体系补充新标签。

**Q: 如何调整分析只看某个市场？**
在 `extract_data.py` 中设置 `MARKET_FILTER = 'US'`（默认 None = 全市场）。

详见：
- `references/taxonomy-design.md` — 三级标签体系设计指南
- `references/report-sections.md` — HTML 报告各模块说明

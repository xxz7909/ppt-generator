# CCTV-17 收视日报 PPT 生成器 — 完整技术规格

> 本文档详尽记录 `generate_report_config_driven.py`（v3 config-driven 生成器）的全部行为与逻辑，
> 同时涵盖它依赖的 `generate_report.py`（草稿生成器）和 `data_reader.py`（数据读取层）。

---

## 1. 总体架构与流程

### 1.1 两阶段管线

v3 生成器采用「**草稿 + 模板回填**」两阶段架构：

```
Excel ──→ data_reader.py ──→ ReportData
                                  │
                                  ├──→ generate_report.py  ──→ 数据草稿 PPT（格式粗糙，数据正确）
                                  │                                  │
                                  │                                  ▼
                                  └──→ generate_report_config_driven.py
                                           │
                                           ├─ 复制 demo 模板 → 输出文件
                                           ├─ XML 级逐页回填草稿数据到 demo 副本
                                           ├─ 修复标题（从 ReportData 直接构建）
                                           ├─ 修复分分钟覆盖表格
                                           ├─ 清除 demo 模板的批注
                                           └─ 保存最终 PPT
```

### 1.2 `generate_report_config_driven()` 主函数流程

```
[1/4] 生成数据稿   → _make_data_draft() → 调用 base.generate_report() 生成临时 .pptx
[2/4] 复制模板      → shutil.copy2(demo, output)
[3/5] XML 级回填    → 打开 target(demo副本) + source(草稿)，逐页调用 _sync_slide()
[4/5] 修复标题      → read_excel_data() 重读 Excel, _fix_titles() 修复 pages 3-8 标题
      修复分分钟页   → _fix_minute_chart_page() 修复第 9/10 页的覆盖表格
[5/5] 保存 + 清注释 → target.save(), _strip_comments()
```

### 1.3 为什么需要两阶段

- **草稿** 由 `generate_report.py` 从零构建，程序化创建图表/表格/文本，数据准确但格式粗糙
- **Demo 模板** 是一份手动制作的参考 PPT，包含完美的字体、字号、颜色、对齐、换行格式
- v3 将草稿数据 **XML 级注入** 到 demo 副本中，保留 demo 的所有 `<a:rPr>`（字体属性）、`<a:br>`（换行）、`<a:pPr>`（段落属性），**只替换 `<a:t>` 文本值**

---

## 2. 命令行参数

### 2.1 v3 `generate_report_config_driven.py`

```
python -X utf8 generate_report_config_driven.py <excel> --template <demo.pptx> --output <output.pptx>
```

| 参数 | 必填 | 说明 |
|------|------|------|
| `excel` | 是 | Excel 数据文件路径 |
| `--template` / `-t` | 是 | Demo 模板 .pptx 路径 |
| `--output` / `-o` | 是 | 输出 PPT 路径 |

### 2.2 草稿生成器 `generate_report.py`

```
python -X utf8 generate_report.py <excel> [--template origin.pptx] [--output report.pptx]
```

| 参数 | 必填 | 说明 |
|------|------|------|
| `excel` | 是 | Excel 数据文件路径 |
| `--template` / `-t` | 否 | PPT 模板（默认 `origin.pptx`）；文件名含"简版"时启用 `BODY_PLAIN_MODE` |
| `--output` / `-o` | 否 | 输出路径（默认 `{basename}_收视日报.pptx`） |

`BODY_PLAIN_MODE`：当模板文件名含"简版"时设为 `True`，跳过标题栏/副标题/页码/装饰元素的生成。v3 调用草稿生成器时传入 demo 模板路径（含"简版"），因此草稿不含标题栏，标题由 v3 的 `_fix_titles()` 单独构建。

---

## 3. 数据流：Excel → ReportData → PPT

### 3.1 Excel 格式

支持两种输入格式：

| 格式 | Sheet 数 | 内容 | 输出页数 |
|------|---------|------|---------|
| 基础（简版） | 6 | 市场份额、电视剧与非电视剧、上星频道排名、台组内排名、分分钟、串单 | 11 页 |
| 完整（日报） | 12 | 基础 6 + 日报-分时段收视率、日报-三大剧场、日报-电视剧观众规模、日报-首播节目对比、日报-首播节目分类观众触达、日报-频道分类观众规模 | 14-17 页 |

### 3.2 数据读取函数 (`data_reader.py`)

`read_excel_data(filepath)` → `ReportData` 是唯一入口，内部按 sheet 名称调度：

| Sheet 名称 | 读取函数 | 数据类 | 关键字段 |
|-----------|---------|--------|---------|
| 市场份额 | `_read_market_share()` | `MarketShareData` | CCTV-17 前推/前一日/当日 收视率/份额/到达率/忠实度 + 所有频道同类指标 + 变化百分比 |
| 市场份额（H-W列） | `_read_audience_from_market_share()` | `List[AudienceItem]` | 16 类分类观众规模（性别2+年龄7+教育5+城乡2），优先级高于日报-频道分类观众规模 |
| 电视剧与非电视剧 | `_read_drama()` | `DramaData` | 电视剧/非电视剧 前推/当日 收视率/份额 |
| 台组内排名 | `_read_org_ranking()` | `List[ChannelRankItem]` | 频道名/简称、前一日份额、当日份额 |
| 上星频道排名 | `_read_channel_ranking()` | `List[ChannelRankItem]` | 同上 + 可选变化列 |
| 串单 | `_read_programs()` | `List[ProgramItem]` | 节目名/集数/日期/星期/开始时间/时长/结束时间/类别/子类/份额/收视率 |
| 串单（首播筛选） | `_read_premiere_from_schedule()` | `List[PremiereProgram]` | 若有"首播"标注列 → 按标注筛选；否则 → 16:30-24:00 非电视剧自动筛选 |
| 分分钟 | `_read_minutes()` | `List[MinuteData]` | 时间/前推收视率/当日收视率/前推份额/当日份额 |
| 日报-分时段收视率 | `_read_time_slots()` | `List[TimeSlotData]` | 时段名/收视率/份额/变化 |
| 日报-三大剧场 | `_read_theaters()` | `List[TheaterData]` | 时段/信息/前推&当日收视率&份额/变化 |
| 日报-首播节目对比 | `_read_premiere_programs()` | `List[PremiereProgram]` | （仅在串单未提供首播数据时使用） |
| 日报-首播节目分类观众触达 | `_read_premiere_audience()` | `Dict` | 各节目各类别的观众触达数据 |
| 日报-频道分类观众规模 | `_read_channel_audience()` | `List[AudienceItem]` | （仅在市场份额H-W列未提供时使用） |
| 日报-电视剧观众规模 | `_read_drama_audience()` | `Dict` | 各时段分类观众 + 各剧场观众规模 |

### 3.3 计算字段

读取完成后自动计算：

- **`_compute_rankings()`**：从 `org_ranking` / `channel_ranking` 中按当日/前日份额排序，定位 CCTV-17 的排名和变化
- **`_compute_premiere_loyalty()`**：忠实度 = 收视率(%) / 到达率(%) × 100；到达率 = 观众规模(万人) / TV总人口(万人) × 100；TV总人口 = 频道到达率(万人) / (频道到达率(%)/100)
- **变化百分比**：`share_change = round((当日 - 前推) / 前推 × 100)`

### 3.4 数据优先级

- **首播节目**：串单 `_read_premiere_from_schedule()` > 日报 `_read_premiere_programs()`
- **分类观众**：市场份额 H-W 列 `_read_audience_from_market_share()` > 日报 `_read_channel_audience()`
- **日期**：文件名 YYYYMMDD > Excel 市场份额中的日期标签 > 兜底"当日"

### 3.5 变化列安全读取

市场份额 sheet 的第 7 列(index 6)起可能是观众规模数据而非变化百分比。代码检查 `rows[0][6]` 表头是否含"变化"/"change"字样，**只有表头明确标注才覆盖计算值**，否则使用计算得到的变化百分比。

---

## 4. 草稿生成器 (`generate_report.py`) — 各 build_* 函数

草稿生成器从零构建 PPT（删除模板所有幻灯片，逐页重建），包含以下构建函数：

| 页码 | 函数 | 内容 |
|------|------|------|
| 1 | `build_cover()` | 封面：背景色/圆形装饰/标题(日期+频道名)/署名(统筹策划部+年份+期号) |
| 2 | `build_summary()` | 收视速报：3段文字 (份额变化/电视剧变化/时段分析) + 菱形项目符号 + 备注 |
| 3 | `build_market_share()` | 市场份额概览：柱状图(份额)+2条形图(所有频道收视率+CCTV-17收视率)+2表格(到达率/忠实度)+描述文本(2段落) |
| 4 | `build_org_ranking()` | 台组内排名：双日期表格(前一日+当日)，CCTV-17高亮 |
| 5 | `build_channel_ranking()` | 上星频道排名：前10+分隔行+CCTV-17上下文行，7列(含变化箭头) |
| 6 | `build_schedule_chart()` | 串单市场份额：组合图(柱=份额, 线=收视率)，过滤6:00-24:00且时长≥10min |
| 7 | `build_program_ranking(metric='rating')` | 栏目收视率排名：前30节目柱状图 |
| 8 | `build_program_ranking(metric='share')` | 栏目收视份额排名：前30节目柱状图 |
| 9 | `build_minute_chart(metric='rating')` | 分分钟收视率：面积图(05:30-25:59)+节目名称标签表+时段汇总表+前月剧集标签表(空1×8) |
| 10 | `build_minute_chart(metric='share')` | 分分钟市场份额：同上 |
| 11 | `build_premiere_comparison(metric='rating')` | 栏目首播收视率：柱状图(前1月均值/前一日/当日或2系列) |
| 12 | `build_premiere_comparison(metric='share')` | 栏目首播市场份额 |
| 13 | `build_channel_audience()` | 频道分类观众规模：19类别含间隔结构的柱状图+变化标签+图例标签 |
| 14 | `build_ending()` | 感谢观看结尾页 |

**简版模式**（`BODY_PLAIN_MODE=True`）跳过：电视剧分析、首播电视剧观众规模、首播忠实度页。

### 4.1 市场份额页的类别结构

```
市场份额柱状图 categories:
  若有前一日: ['前1个月\n均值', '2月14日', '2月15日']  (3 个)
  若无前一日: ['前1个月\n均值', '2月15日']             (2 个)
```

描述文本刻意生成 **2 个独立 `<a:p>` 段落**（而非 `\n` 连接），确保 v3 sync 时与 demo 的 2 段落结构匹配，不留残留文字。

### 4.2 分分钟页的 3 表结构

每个分分钟 slide 包含：
1. **节目名称标签表** (1×N)：按串单节目的开始/结束时间分配列宽
2. **时段汇总表** (1×6/N)：各时段的当日值 + 变化百分比
3. **前月剧集标签表** (1×8)：空占位符，与 demo 结构一致

---

## 5. Demo 模板复制与 XML 级同步机制

### 5.1 核心思路

**不修改任何格式属性，只替换 `<a:t>` 文本内容。**

```xml
<!-- demo 原始 -->
<a:r>
  <a:rPr lang="zh-CN" sz="1400" b="1">
    <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
  </a:rPr>
  <a:t>0.210</a:t>   ← 只替换这里
</a:r>
```

### 5.2 XML 命名空间

```python
A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'    # <a:...> DrawingML
P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main' # <p:...> PresentationML
C_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart'   # <c:...> Chart
```

- **形状文本体**: `<p:txBody>` 或 `<a:txBody>`
- **表格单元格文本体**: `<a:txBody>`（在 `<a:tc>` 内）

### 5.3 文本操作函数

| 函数 | 作用 |
|------|------|
| `_get_para_text(p_elem)` | 从 `<a:p>` 提取纯文本，`<a:br>` → `\n` |
| `_get_runs(p_elem)` | 获取段落内所有 `<a:r>` 元素 |
| `_set_run_text(run_elem, text)` | 设置 run 的 `<a:t>`，保留 `<a:rPr>` |
| `_replace_para_text(target_p, new_text)` | **核心**: 若有 `<a:br>` 按 `\n` 分段分配到 run 组；否则全部写入首个 run，清空其余 |
| `_clear_para_text(target_p)` | 清空段落所有 run 文本 |

### 5.4 `_replace_para_text` 的分段策略

**无 `<a:br>` 时**：
- 所有文本写入第一个 `<a:r>` 的 `<a:t>`
- 其余 `<a:r>` 的 `<a:t>` 设为空字符串
- 若源文本含 `\n`（来自 `<a:br>`），只取第一行

**有 `<a:br>` 时**：
- 按 `\n` 拆分新文本为 segments
- 按 XML 子元素顺序将 run 分组（以 `<a:br>` 为分隔）
- 每组的第一个 run 写入对应 segment，其余 run 清空

### 5.5 控制字符清洗

```python
_VTAB_PATTERN = re.compile(r'_x000[0-9A-Fa-f]_')

def _sanitize(text):
    # 清除 _x000B_ 等 OOXML 转义和 \x0b (VT) / \x0c (FF) 控制字符
```

### 5.6 形状匹配策略

**两轮匹配** (`_match_nearest`)：

1. **内容匹配**（短文本 ≤ 30字符）：防止静态标签（如"前1个月均值"）被交叉覆盖
2. **最近邻位置匹配**：按形状中心点距离匹配剩余形状

**图表匹配** (`_match_charts_by_y_order`)：按 `top` Y 坐标排序配对，确保上方图表匹配上方图表、下方匹配下方（最近邻在位置偏差大时会交叉）。

### 5.7 文本同步 (`_sync_text_shape`)

**非空段落对齐策略**：
1. 提取 target/source 各自的非空段落索引
2. 按顺序一一配对（空段落=空行分隔符，保持不变）
3. 内容相同时跳过（统一 `\x0b`/`\n` 后比较）
4. 源段落多于目标时，追加到最后一个已配对的目标段落（`\n` 连接）

---

## 6. `_sync_slide` — 逐页同步分发

```python
def _sync_slide(target_slide, src_slide, slide_index, total_slides):
```

| slide_index | 页面 | 处理方式 |
|-------------|------|---------|
| 0 | 封面 | `_update_cover()` 精确替换日期/期号 |
| n-1 | 感谢观看 | 跳过（不做任何替换） |
| 12 | 频道分类观众规模 | `_sync_and_position_audience()` 专用 |
| 5 | 串单市场份额 | 文本同步 + `_fix_schedule_chart_page()` 组合图表 + 超轴标注 |
| 8, 9 | 分分钟页 | 文本同步 + 图表同步，**跳过通用表格同步**（由 `_fix_minute_chart_page` 处理） |
| 其他 | 通用 | 文本+图表+表格 三类形状分别匹配同步 |

**通用同步后的修正 (post-fix)**：

| slide_index | 修正函数 | 作用 |
|-------------|---------|------|
| 2 | `_fix_change_pct_colors()` | 表格变化百分比: 正值→红，负值→绿 |
| 2 | `_fix_page3_numfmt()` | 所有图表数据标签 → 3 位小数 |
| 3 | `_fix_org_ranking_page()` | CCTV-17行红色+粗体, 橘色阈值线定位 |
| 4 | `_fix_channel_ranking_page()` | ↑红↓黑, CCTV-17行全红+粗体 |
| 6 | `_fix_chart_max_annotation()` | 超轴柱标注，格式 `:.4f` |
| 7 | `_fix_chart_max_annotation()` | 超轴柱标注，格式 `:.3f` |
| 10 | `_fix_premiere_chart_axis()` + `_fix_premiere_chart_colors(橙)` | 首播收视率页轴+颜色 |
| 11 | `_fix_premiere_chart_axis()` + `_fix_premiere_chart_colors(绿)` | 首播市场份额页轴+颜色 |

---

## 7. 封面处理 (`_update_cover`)

精确更新 demo 封面的日期和期号，**不改变任何格式**：

**标题 shape**（含"频道收视日报"/"农业"）：
- 逐 run 扫描，识别年/月/日数字：
  - 4 位数字 → 年份
  - 1-2 位数字 + 后一个 run 含"月" → 月份
  - 1-2 位数字 + 前一个 run 含"月" → 日期

**署名 shape**（含"统筹策划部"/"策划部"）：
- 匹配 `\s*\d{4}` 格式 → 替换年份（保留前导空白）
- 纯数字 + 前一个 run 含"第" → 替换期号

---

## 8. 标题生成 (`_fix_titles` + `_build_slide_titles`)

简版模式下草稿不含标题栏，v3 需要**直接从 ReportData 构建标题文本**并写入 demo 中已有的标题 shape。

```python
_build_slide_titles(data) → {slide_index: title_text}
```

| 页码 (idx) | 标题模板 |
|------------|---------|
| 3 (idx=2) | `{日期}，市场份额{变化描述}，收视率{变化描述}` |
| 4 (idx=3) | `央视台组排名第{N}位，较前一日{升降描述}` |
| 5 (idx=4) | `上星频道排名第{N}位，较前一日{升降描述}` |
| 6 (idx=5) | `{日期}，频道市场份额{x.xxx}%，收视率{x.xxx}%` |
| 7 (idx=6) | `{日期}栏目收视率排名` |
| 8 (idx=7) | `{日期}栏目收视份额排名` |

变化描述逻辑 (`_change_desc`)：|pct| ≤ 5 → "基本持平"；> 0 → "提升N%"；< 0 → "下降N%"

标题查找 (`_find_title_shape`)：遍历 slide.shapes，找 `name` 含"标题"的文本框。

---

## 9. 图表处理

### 9.1 通用图表同步 (`_sync_chart`)

```
1. 保存 demo 模板的 dLbls/numFmt（replace_data 会重置）
2. chart.replace_data(chart_data)          ← python-pptx API
3. 恢复保存的 numFmt
4. 若模板无 numFmt，尝试从 draft 复制
```

### 9.2 组合图表同步 (`_sync_combo_chart`)

用于第 6 页串单市场份额页（barChart + lineChart 双图结构）。

**不使用 `replace_data`**（会破坏双图结构），而是 XML 级替换：
- 从 src 提取：categories, bar_values (series 0=市场份额), line_values (series 1=收视率)
- 在 target 的 `<c:barChart><c:ser>` 和 `<c:lineChart><c:ser>` 中分别替换 `strCache`（分类）和 `numCache`（数值）
- 更新 `ptCount` 以匹配新数据长度

### 9.3 图表数据提取 (`_extract_chart_data`)

从 python-pptx chart 对象提取 `CategoryChartData`：categories + 各 series 的 name 和 values。

### 9.4 坐标轴计算 (`_compute_nice_axis`)

模拟 PowerPoint 自动缩放：
```python
raw_interval = max_val / 6
magnitude = 10 ** floor(log10(raw_interval))
candidates = [1x, 2x, 5x, 10x] × magnitude
选择使 axis_max 最小且 tick 数在 5-12 之间的 interval
axis_max = ceil(max_val / interval + 1) × interval
```

### 9.5 超轴标注 (`_fix_chart_max_annotation`)

当柱子超过 Y 轴最大值时，在对应位置放置数值标注文本框：
1. 从 barChart XML 找出所有 `value > axis_max` 的数据点
2. 收集页面上名为"文本框 N"的候选标注框，按 left 排序
3. 按 bar index 对应分配，用柱位置反推 X 坐标，按溢出比例计算 Y
4. 若无 run 存在则从同页其他标注文本框复制格式（deepcopy `<a:r>`）

### 9.6 分类轴字号自适应 (`_auto_fit_catax_font`)

第 6 页串单图：基准 16 字符 → 14pt (sz=1400)，标签更长时按比例缩小到最小 6.5pt (sz=650)。

---

## 10. 表格同步 (`_sync_table`)

逐单元格 XML 级替换：
- 取 min(target_rows, src_rows)、min(target_cols, src_cols)
- 每个单元格：找 `<a:txBody>` → 取第一个 `<a:p>` → `_replace_para_text()`
- 多余段落清空

---

## 11. Page 13 — 频道分类观众规模 (`_sync_and_position_audience`)

这是最复杂的单页处理，包含 7 个步骤：

### 11.1 图表数据同步
调用通用 `_sync_chart()` 替换数据。

### 11.2 清除 "None" 间隔类别
`replace_data` 会保留间隔类别的 `None` 文本。在 XML 中删除 `strCache` 里 `<v>` 为 "None" 或空的 `<pt>` 元素。

### 11.3 图表几何计算
从 `<c:plotArea><c:layout><c:manualLayout>` 读取 x/y/w/h 百分比，换算为 EMU 绝对坐标：
```
plot_left   = chart.left + x% × chart.width
plot_top    = chart.top  + y% × chart.height
plot_width  = w% × chart.width
plot_height = h% × chart.height
plot_bottom = plot_top + plot_height
```

### 11.4 读取图表数据 & 计算变化百分比
- series[0] = 前期，series[1] = 当日
- 跳过 `max_v ≤ 0` 的间隔类别
- 变化：`(当日 - 前期) / 前期 × 100`
- 格式：|val| < 1% 保留 1 位小数，否则取整；负值加 `-` 前缀

### 11.5 计算坐标轴范围 & 设置显式最大值
用 `_compute_nice_axis()` 计算，写入 `<c:valAx><c:scaling><c:max>`。

### 11.6 分离柱上标签 vs 图例标签
- 以 `↑↓▲▼` 开头 → **图例标签**（如 `↓8%`，总体变化）
- 以数字 / `-` 开头 → **柱上标签**（各类别变化）

### 11.7 定位每个柱上标签
```
X = plot_left + plot_width × (cat_index + 0.5) / num_categories - label_width / 2
Y = plot_bottom - (max_value / axis_max) × plot_height - label_height - GAP_EMU(55000)
```
颜色：红升绿降 (`_set_label_color`)

### 11.8 图例变化标签
计算 total_s0/total_s1 的总变化百分比，格式 `↑N%` / `↓N%`，颜色同上。

### 11.9 19 类别结构

| 组 | 类别 | 数量 |
|----|------|------|
| 性别 | 男、女 | 2 |
| *间隔* | 空 | 1 |
| 年龄 | 4-14岁, 15-24, 25-34, 35-44, 45-54, 55-64, 65+ | 7 |
| *间隔* | 空 | 1 |
| 教育 | 未受过正规教育, 小学, 初中, 高中, 大学以上 | 5 |
| *间隔* | 空 | 1 |
| 城乡 | 城市, 乡村 | 2 |

共 16 类别 + 3 个空间隔 = 19 个 chart categories。

---

## 12. 分分钟页修正 (`_fix_minute_chart_page`)

修正第 9/10 页 (slide_index=8,9) 的 3 个覆盖表格：

### 12.1 表格识别
- **表头表格** (8 cols, 有文字)：固定时段名称
- **分隔线表格** (8 cols, 空)：绿色竖线对齐
- **节目名称表格** (15+ cols)：串单节目名

### 12.2 表头更新
固定 8 个时段：早间节目(06:00-10:00)、上午剧场(10:00-12:00)、午间节目(12:00-13:30)、下午剧场(13:30-17:00)、傍晚节目(17:00-18:30)、黄金剧场(18:30-20:30)、晚间节目(20:30-22:30)、夜间节目(22:30-25:00)

### 12.3 节目名称表格更新
1. 用 `_group_programs()` 将串单节目合并为 ≤35 组
2. `_resize_prog_table()` 动态调整列数（通过克隆/删除最后一列的 gridCol 和 tc 实现）
3. `_proportional_widths()` 按时长比例分配列宽
4. 对齐到图表绘图区（从 chart XML 解析 plotArea x/w）
5. 节目段连续覆盖整个图表时间范围（首段起点=chart起点，末段终点=chart终点）
6. 首播节目着色：收视率页=橙色(F39C12)，市场份额页=绿色(4A7C31)

---

## 13. `_group_programs` — 节目分组逻辑

将串单节目合并为可在表格中展示的组：

1. **过滤**：排除名称为{国歌,歌曲,再见}、凌晨 02:00-05:29 的节目
2. **首播/重播判定**：13:30-22:30 → 首播，其余 → 重播，添加对应后缀
3. **连续合并**：同名+同首/重播状态的连续节目合并（如电视剧连续几集）
4. **短节目合并**：时长 < 10min 的节目与后一个节目合并（名称用 `+` 连接，长度上限 20 字符）
5. **超限合并**：若组数仍 > max_cols(35)，反复合并最短的相邻组
6. **截断**：合并后仍超长的名称只保留第一个节目名

---

## 14. 其他修正函数

### 14.1 `_fix_org_ranking_page` — 台组排名页 (Page 4)

1. 清除 demo 模板遗留的红色单元格边框 (FF0000 tcBdr)
2. 所有数据行恢复为黑色、非粗体
3. CCTV-17 所在**半区**（左3列或右3列）设为红色+粗体（避免对侧误染）
4. 橘色阈值线定位：从表格 XML 累加行高 (`<a:tr h="...">`) 找到右侧份额首次 < 阈值的行，更新连接符（直接连接符 10）Y 坐标
5. 阈值标签文本恢复（text sync 可能覆盖了 demo 的值）

阈值配置从 `demo_layout_config.json` 读取（`_load_threshold_config`）：slideNo=4 的"文本框 10"的颜色和文本。

### 14.2 `_fix_channel_ranking_page` — 上星频道排名页 (Page 5)

1. 箭头列 (col 6)：`↑` → 红色(FF0000), `↓` → 黑色(000000)
2. CCTV-17 行：全行红色 + 粗体

### 14.3 `_fix_change_pct_colors` — 市场份额页表格颜色 (Page 3)

匹配 `[+-]?\d+\.?\d*%` 格式的单元格，正值→红(FF0000)，负值→绿(00B050)。

### 14.4 `_fix_page3_numfmt` — 市场份额页数据标签格式 (Page 3)

将所有图表 dLbls 的 numFmt 从 2 位小数 (`0.00`) 统一为 3 位 (`0.000`)。

### 14.5 `_fix_premiere_chart_axis` — 首播页 Y 轴 (Pages 11-12)

demo 模板有固定 Y 轴上限，replace_data 后若新数据超出则用 `_compute_nice_axis` 重新计算。
若数据远小于当前上限 (< 50%)，也调低避免留白过大。

### 14.6 `_fix_premiere_chart_colors` — 首播页系列颜色 (Pages 11-12)

当只有 2 个系列时（无月均数据），最后一个系列（当日）继承了 demo 的浅灰色，需改为指定颜色（收视率=橙 FE9B1C，份额=绿 4A7C31）。通过 XML 修改 `<c:ser><c:spPr><a:solidFill>` 实现。

### 14.7 `_strip_comments` — 清除批注

ZIP 级别删除 `commentAuthors.xml` / `comments/*.xml`，清理 `[Content_Types].xml` 和 `.rels` 中的引用。替换原文件。

---

## 15. 特殊文本处理

### 15.1 Unicode / 控制字符规范化

| 场景 | 处理方式 |
|------|---------|
| `_x000B_` 等 OOXML 控制字符转义 | `_VTAB_PATTERN` 正则移除 |
| `\x0b` (vertical tab) | `.replace('\x0b', '')` |
| `\x0c` (form feed) | `.replace('\x0c', '')` |
| `\x0b` 在比较时 | 统一替换为空/`\n` 后比较 |

### 15.2 换行符处理

| 来源 | 格式 | 处理 |
|------|------|------|
| Demo 模板 | `<a:br>` 元素 → python-pptx 读取为 `\x0b` | 保留原始 `<a:br>` XML 结构 |
| 草稿生成器 | 段落分隔 `\n`（多个 `<a:p>`） | 按非空段落对齐 |
| 文本同步 | `_get_para_text` 返回中 `<a:br>` → `\n` | `_replace_para_text` 按 `\n` 分段写回 |

### 15.3 `_sanitize` 函数

```python
def _sanitize(text):
    s = str(text)
    s = re.sub(r'_x000[0-9A-Fa-f]_', '', s)  # 清除 OOXML 控制字符转义
    s = s.replace('\x0b', '').replace('\x0c', '')  # 清除 VT / FF
    return s
```

---

## 16. `_set_label_color` / `_set_cell_font_color` — 颜色修改

### shape 级颜色修改 (`_set_label_color`)
遍历 `<a:rPr>` 和 `<a:endParaRPr>`，移除旧 `<a:solidFill>`，添加新 `<a:srgbClr val="RRGGBB">`。

### 表格单元格级颜色修改 (`_set_cell_font_color`)
同上，额外处理：
- `<a:defRPr>`（顶层默认运行属性）
- `<a:pPr>/<a:defRPr>`（段落级默认运行属性）
- `<a:endParaRPr>`（段落结尾）
- 可选设置 `b="0"` / `b="1"` 粗体

---

## 17. 数据类结构总览

```python
@dataclass
class ReportData:
    # 元数据
    report_date: str              # "2026年2月15日"
    report_date_short: str        # "2月15日"
    period_label: str             # "2026/1/11-2026/2/9"

    # 基础数据
    market_share: MarketShareData      # 市场份额（CCTV-17 + 所有频道 × 前推/前一日/当日）
    drama: DramaData                   # 电视剧/非电视剧 收视率/份额
    org_ranking: List[ChannelRankItem] # 台组内排名
    channel_ranking: List[ChannelRankItem] # 上星频道排名
    programs: List[ProgramItem]        # 串单（完整节目表）
    minutes: List[MinuteData]          # 分分钟数据

    # 扩展数据（日报）
    has_daily_report: bool
    time_slots: List[TimeSlotData]
    theaters: List[TheaterData]
    premiere_programs: List[PremiereProgram]
    premiere_audience: Dict
    channel_audience: List[AudienceItem]
    drama_audience: Dict

    # 计算字段
    org_rank: int / org_rank_change: int
    channel_rank: int / channel_rank_change: int
```

---

## 18. 输出 PPT 页面结构（简版 14 页）

| 页码 | slide_index | 内容 | 同步方式 |
|------|-------------|------|---------|
| 1 | 0 | 封面 | `_update_cover` 精确日期替换 |
| 2 | 1 | 收视速报 | 通用文本同步 |
| 3 | 2 | 市场份额概览 | 通用 + `_fix_change_pct_colors` + `_fix_page3_numfmt` |
| 4 | 3 | 台组内排名 | 通用 + `_fix_org_ranking_page` |
| 5 | 4 | 上星频道排名 | 通用 + `_fix_channel_ranking_page` |
| 6 | 5 | 串单市场份额 | 通用文本 + `_fix_schedule_chart_page`(combo chart) |
| 7 | 6 | 栏目收视率排名 | 通用 + `_fix_chart_max_annotation(:.4f)` |
| 8 | 7 | 栏目收视份额排名 | 通用 + `_fix_chart_max_annotation(:.3f)` |
| 9 | 8 | 分分钟收视率 | 通用文本/图表 + `_fix_minute_chart_page(rating)` |
| 10 | 9 | 分分钟市场份额 | 通用文本/图表 + `_fix_minute_chart_page(share)` |
| 11 | 10 | 栏目首播收视率 | 通用 + `_fix_premiere_chart_axis` + `_fix_premiere_chart_colors(橙)` |
| 12 | 11 | 栏目首播市场份额 | 通用 + `_fix_premiere_chart_axis` + `_fix_premiere_chart_colors(绿)` |
| 13 | 12 | 频道分类观众规模 | `_sync_and_position_audience` 专用 |
| 14 | 13 | 感谢观看 | 跳过 |

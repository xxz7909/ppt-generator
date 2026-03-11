# 当前任务

> CCTV-17 收视日报 PPT 生成器 - 任务跟踪

## 进行中

- [ ] (无)

## 待办
- [x] 分分钟Page9/10节目标签变色：改为读取串单P列（首重播）实际数据，仅"首播"节目变色
  - ProgramItem新增premiere_type字段
  - _read_programs读取P列(col 15)
  - _group_programs由时间段推断改为按premiere_type=='首播'判定
- [x] Page2速报文本修复："与前一日" → "较前一日"；排名变化用词"上升" → "提升"
- [x] 分分钟Page9/10分时段标签严格对齐图表横坐标
  - 表头表格/分隔线表格保持原始尺寸，列宽按图表时段边界位置计算
  - 列分隔线精确对齐 10:00/12:00/13:30/17:00/18:30/20:30/22:30 刻度
  - 修复z-order：节目表格始终在分隔线表格之上，P10节目名称可编辑
- [ ] 打包成GUI-exe并制作友好的交互与记忆功能

## 已完成

- [x] Page13频道分类观众规模：动态定位变化百分比标签（紧贴柱上方）
  - 19类别结构（含3个间隔Gap），匹配demo布局
  - 从图表数据动态计算变化百分比（非文本匹配）
  - 基于ManualLayout计算PlotArea几何 → bar_top → label_top
  - 自动nice axis算法（_compute_nice_axis）
  - XML清除replace_data产生的"None"间隔类别
  - 箭头前缀区分legend标签 vs bar标签

- [x] 市场份额页3类别图表(前1月均值/前一日/当日)，匹配demo结构
- [x] 市场份额页描述文本双段落(消除demo残留文本)
- [x] 修复变化列误读bug(观众规模列被当作变化百分比)
- [x] 分分钟slide添加"前一个月剧集"空标签表(1r×8c)
- [x] 市场份额页收视率柱状图替换(reach/loyalty → all_rating/cctv17_rating匹配demo)

- [x] 项目架构搭建(ppt_config.py, data_reader.py, slide_utils.py, generate_report.py)
- [x] Excel数据读取(支持6-sheet和12-sheet两种格式)
- [x] 17种幻灯片构建(封面/速报/市场份额/排名/串单/节目排名/分分钟/电视剧/观众规模/首播对比/忠实度/结尾)
- [x] 电视剧观众规模双图表修复(slide 12)
- [x] 忠实度计算(从rating/audience/universe推导，替代错误的share值)
- [x] 频道分类观众规模变化标签(↑/↓箭头)
- [x] 首播节目类别标签添加时段信息
- [x] 摘要文本"基本持平"阈值修正(≤5%)
- [x] 双格式测试通过(17页/11页)
- [x] v3 config-driven生成器: XML级别复刻demo格式(generate_report_config_driven.py)
- [x] v3 命名空间修复(p:txBody vs a:txBody)
- [x] v3 标题补全系统(_fix_titles + _build_slide_titles)
- [x] v3 非空段落对齐策略(跳过empty separator paragraphs)
- [x] v3 两轮匹配(content-first防标签错位 + position-based)
- [x] v3 \x0b/\n规范化
- [x] v3 Page 13专用handler(X坐标排序匹配)
- [x] v3 批量生成3份PPT并通过验证(20260215/20260202/20260128)

---

## 使用说明

### 任务状态
- `- [ ]` 待办/进行中
- `- [x]` 已完成

### 分类
- **进行中**: 当前正在处理的任务
- **待办**: 计划要做但还没开始的任务
- **已完成**: 已经完成的任务

### 更新方式
1. Claude 会在工作时自动更新这个文件
2. 你也可以直接编辑这个文件
3. 下次会话时，Claude 会读取这个文件来了解任务状态

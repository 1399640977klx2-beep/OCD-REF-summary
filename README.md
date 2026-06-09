# OCD Toolbox

半导体 OCD 量测建模辅助工具。PyQt5 桌面应用，四个标签页覆盖 REF 汇总、Match 匹配、Wafer 地图绘制、光谱文件夹管理。

## 运行环境

- Python 3.13 (conda: `loadref`)
- PyQt5 + pandas + matplotlib + openpyxl + scipy

## 启动

```bash
conda activate loadref
python main.py
```

## 项目结构

```
OCD Toolbox/
├── main.py                       # 入口
├── ui/                           # 界面层
│   ├── main_window.py            # 主窗口 + 4 个 Tab
│   ├── tab_ref_aggregator.py     # Tab 1: REF 数据汇总
│   ├── tab_match.py              # Tab 2: Match 匹配
│   ├── tab_wafer_map.py          # Tab 3: Wafer 地图
│   └── tab_folder_manager.py     # Tab 4: 文件夹管理
├── parsers/                      # 数据解析层
│   ├── __init__.py               # 调度器 (parse_files)
│   ├── kla_parser.py             # KLA 格式
│   ├── nova_parser.py            # NOVA 格式
│   └── pmish_parser.py           # PMISH 格式
├── utils/                        # 工具层
│   ├── file_utils.py             # 文件扫描、WaferID 提取
│   └── excel_writer.py           # Excel 输出
└── data/                         # 测试数据 (不在 git 中)
```

## 四个标签页功能

### Tab 1: REF 数据汇总
- 选择厂商 (KLA/NOVA/PMISH) → 选择数据文件夹 → 解析 → 导出汇总 Excel
- PMISH 数据自动按 PAD 名称分 Sheet 输出
- 自动跳过统计摘要行 (MAX/MIN/AVERAGE 等)

### Tab 2: Match 匹配
- 加载公司软件输出 Excel + BSL/REF 数据 → 匹配 → 导出 Match Excel
- 匹配键: WaferID + FIELD X/FIELD Y ↔ WaferID + Col/X/Row/Y
- 支持选择匹配参数 (DP, HIGH 等) 后再生成 Mean&STD 统计

### Tab 3: Wafer 地图
- 从原 `draw_wafer_map.py` 迁移，功能完整保留
- 加载 Excel → 选 Sheet → 选 Wafer/参数 → contour map 绘制 + 批量保存

### Tab 4: 文件夹管理
- 整合 `renamefolder.py` 和 `re_renamefolder.py`
- 自动重命名 (timestamp → WaferID_timestamp) + 手动修正前缀 + 批量替换

## Parser 实测

| 厂商 | 数据源 | 行数 | 列数 | 说明 |
|------|--------|------|------|------|
| KLA | SGDEP | 696 | 30 | 跳过统计 + 元数据提取 |
| NOVA | DGALL + SP1AEI | 3332+ | 41 | `index_col=False` 修复错列 |
| PMISH | DGALL + SGDEP | 852+78 | 50-57 | 多段拼合 + 括号列名统一 + 按 Pad 分 Sheet |

## 已知问题 / 待处理

- **Tab 2 匹配逻辑**: 框架已搭好，但具体匹配逻辑 (WaferID 映射不同 Waver 名、DieSeq vs DieOrder 等) 需要用户在实际数据上验证后逐步细化
- **KLA 解析**: 基础解析可用，但未在所有 KLA 文件上充分测试
- **Wafer Map Tab**: 中文字符显示可能受系统编码影响

## Bug 修复记录

| 日期 | 问题 | 修复 |
|------|------|------|
| 06-08 | NOVA 数据列左移 (WaferID→ToolName) | `index_col=False` 防止 pandas 按尾逗号误建索引 |
| 06-08 | Parser 静默失败 | `parse_files()` 新增 `error_handler` 回显错误到界面 |
| 06-08 | PMISH Wafer/Lot 列为空 | 重写 header 元数据解析 (`key:,value` 格式) |
| 06-08 | PMISH 只读第一段 (多段结构) | 改用 `io.StringIO` 动态拼接所有数据段 |
| 06-08 | PMISH 列名括号导致重复列 | `re.sub(r'\([^)]*\)', '', col)` 统一去括号 |
| 06-08 | PMISH 多 PAD 混一 Sheet | `write_ref_summary_by_pad()` 按 PAD 名分 Sheet |

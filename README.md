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
│   ├── __init__.py               # 调度器 (parse_files, 支持 dict/DataFrame 返回)
│   ├── kla_parser.py             # KLA 格式 (Site # 锚点方案, 返回 dict)
│   ├── nova_parser.py            # NOVA 格式 (index_col=False 防错列)
│   └── pmish_parser.py           # PMISH 格式 (多段拼合 + StringIO)
├── utils/                        # 工具层
│   ├── file_utils.py             # 文件扫描、WaferID 提取
│   └── excel_writer.py           # Excel 输出 (含多 Sheet 支持)
└── data/                         # 测试数据
```

## 四个标签页功能

### Tab 1: REF 数据汇总 ✅ 完成
- 选择厂商 (KLA/NOVA/PMISH) → 选择数据文件夹 → 解析 → 导出汇总 Excel
- 自动跳过统计摘要行 (MAX/MIN/AVERAGE 等)
- PMISH / KLA 数据自动按 PAD 名称分 Sheet 输出
- KLA 采用 `Site #` 锚点方案：向上扫元数据、向下切数据，天然保证列序正确

### Tab 2: Match 匹配 🔶 框架完成，逻辑待细化
- 加载公司软件输出 Excel + BSL/REF 数据 → 匹配 → 导出 Match Excel
- 匹配键: WaferID + FIELD X/FIELD Y ↔ WaferID + Col/X/Row/Y
- 支持选择匹配参数 (DP, HIGH 等) 后再生成 Mean&STD 统计

### Tab 3: Wafer 地图 🔶 框架迁移完成
- 从原 `draw_wafer_map.py` 迁移，功能完整保留
- 加载 Excel → 选 Sheet → 选 Wafer/参数 → contour map 绘制 + 批量保存

### Tab 4: 文件夹管理 ✅ 完成
- 整合 `renamefolder.py` 和 `re_renamefolder.py`
- 自动重命名 (timestamp → WaferID_timestamp) + 手动修正前缀 + 批量替换

## Parser 实测

| 厂商 | 数据源 | 行数 | 列数 | 技术要点 |
|------|--------|------|------|----------|
| KLA | SGDEP (4 文件) | 416 | 29-31 | `Site #` 锚点 + 向上扫元数据 + 边界检测 + 返回 dict |
| NOVA | DGALL + SP1AEI | 3332+ | 41 | `index_col=False` 修复尾逗号错列 + 去重表头 |
| PMISH | DGALL + SGDEP | 930+ | 50-57 | 多段 `io.StringIO` 拼合 + 括号列名归一 + 按 Pad 分 Sheet |

## 已知问题 / 待处理

- **Tab 2 匹配逻辑**: 框架已搭好，具体匹配逻辑 (WaferID 映射、DieSeq vs DieOrder 等) 待用户实际验证后细化
- **Tab 3 Wafer Map**: 未充分测试实际数据导入和图像输出
- **KLA Parser**: 仅验证 SGDEP 数据，其他 Layer 格式可能存在差异

## Bug 修复记录

| 问题 | 修复 |
|------|------|
| NOVA 数据列左移 (WaferID→ToolName) | `index_col=False` 防止 pandas 按尾逗号误建索引 |
| Parser 静默失败 | `parse_files()` 新增 `error_handler` 回显错误到界面 |
| NOVA 尾逗号多一列 | 去 `pd.read_csv` 的 `encoding='utf-8'` |
| PMISH Wafer/Lot 列为空 | 重写 header 元数据解析 (`key:,value` 格式) |
| PMISH 只读第一段 (多段结构) | 改用 `io.StringIO` 动态拼接所有数据段 |
| PMISH 列名括号导致重复列 | `re.sub(r'\([^)]*\)', '', col)` 统一去括号 |
| PMISH 多 PAD 混一 Sheet | `write_ref_summary_by_pad()` 按 PAD 名分 Sheet |
| KLA Parser 重构 (旧方案) | 状态机 → `Site #` 锚点方案，代码从 80 行减到 55 行 |
| KLA 元数据全空 | TEST LABEL 后的 `break` 误中止向上扫描 |
| KLA PAD 混在一起 | 向上扫描跨入上一段数据区 → 加边界检测 (遇数据行停) |
| KLA 导出闪退 | `import pandas as pd` 缺失 |
| KLA WaferID/Lot/Recipe 列位置 | `df.insert(0, ...)` 替代 `df['col'] =` 放最左侧 |
| PMISH/KLA 列跨 PAD 空值 | 导出时 `dropna(axis=1, how='all')` 去掉全空列 |

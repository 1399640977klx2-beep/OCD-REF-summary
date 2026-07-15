# OCD Toolbox

半导体 OCD 量测建模辅助工具。PyQt5 桌面应用，六个标签页覆盖 REF 汇总、Match 匹配、Wafer 地图、文件夹管理、光谱整理。

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
│   ├── main_window.py            # 主窗口 + 6 个 Tab
│   ├── tab_ref_aggregator.py     # Tab 1: REF 数据汇总
│   ├── tab_match.py              # Tab 2: Match 匹配
│   ├── tab_wafer_map.py          # Tab 3: Wafer 地图（按钮 → 原版窗口）
│   ├── tab_folder_manager.py     # Tab 4: 文件夹管理
│   ├── tab_organize_spectra.py   # Tab 5: 光谱整理
│   └── draw_wafer_map.py         # Tab 3 引用的原版 Wafer 地图代码
├── parsers/                      # 数据解析层
│   ├── __init__.py               # 调度器 (parse_files, 支持 dict/DataFrame 返回)
│   ├── kla_parser.py             # KLA 格式 (Site # 锚点方案, 返回 dict)
│   ├── nova_parser.py            # NOVA 格式 (index_col=False 防错列)
│   └── pmish_parser.py           # PMISH 格式 (多段拼合 + StringIO)
├── utils/                        # 工具层
│   ├── file_utils.py             # 文件扫描、WaferID 提取
│   ├── excel_writer.py           # Excel 输出 (含多 Sheet 支持)
│   └── excel_writer_match.py     # Match Sheet 生成 (openpyxl 公式/布局)
└── data/                         # 测试数据
```

## 六个标签页功能

### Tab 1: REF 数据汇总 ✅
- 选择厂商 (KLA/NOVA/PMISH) → 选择数据文件夹 → 解析 → 导出汇总 Excel
- 自动跳过统计摘要行 (MAX/MIN/AVERAGE 等)
- PMISH / KLA 数据自动按 PAD 名称分 Sheet 输出
- KLA 采用 `Site #` 锚点方案
- Parser 坐标列名支持可配置列表 (`X_COLS`, `Y_COLS`)

### Tab 2: Match 匹配 ✅
- 加载 Excel → 手动选择 BSL Sheet 和公司数据 Sheet
- Scroll 区域双列网格展示 BSL 全部表头，选公司 Sheet 后红色高亮缺失参数
- 生成 Match Sheet（全公式驱动布局）：
  - BSL 参考数据 + 层级轴标签 + 公司识别 + 原始/校准/REF/Bias 列
  - 校准系统：Row 1-2 存 SLOPE/INTERCEPT，$行号绝对引用
  - 支持动态参数数 `n`，列位置自动计算
  - BSL 坐标列自动检测三种格式

### Tab 3: Wafer 地图 ✅
- 按钮 → 打开原版 `draw_wafer_map.py` 独立窗口
- 所有原版功能完整保留（加载 Excel → 选 Sheet → contour map 绘制 + 批量保存）

### Tab 4: 文件夹管理 ✅
- 整合 `renamefolder.py` 和 `re_renamefolder.py`
- 自动重命名 (timestamp → WaferID_timestamp) + 手动修正前缀 + 批量替换

### Tab 5: 光谱整理 ✅
- 封装原版 `organize_spectra.py` 的 `organize()` 函数
- 源/目标目录选择 + 产品正则 + 机台正则 + Dry Run 预览
- 执行日志实时显示在界面底部的文本区域

## Parser 实测

| 厂商 | 数据源 | 行数 | 列数 | 技术要点 |
|------|--------|------|------|----------|
| KLA | SGDEP (4 文件) | 416 | 29-31 | `Site #` 锚点 + 向上扫元数据 + 边界检测 + 返回 dict |
| NOVA | DGALL + SP1AEI | 3332+ | 41 | `index_col=False` 修复尾逗号错列 + 去重表头 |
| PMISH | DGALL + SGDEP | 930+ | 50-57 | 多段 `io.StringIO` 拼合 + 括号列名归一 + 按 Pad 分 Sheet |

## Bug 修复记录

| 问题 | 修复 |
|------|------|
| NOVA 数据列左移 | `index_col=False` |
| PMISH Wafer/Lot 列为空 | 重写 header 元数据解析 |
| PMISH 多段只读第一段 | `io.StringIO` 动态拼接 |
| PMISH 列名括号重复 | `re.sub` 统一去括号 |
| KLA 元数据全空 | 去 break 加边界检测 |
| KLA PAD 混在一起 | 向上扫描遇数据行停 |
| Step2 参数不显示 | `takeAt(0)` 替换 `deleteLater()` |
| Match 公式列硬编码 | `mid_start` 动态计算 |
| Bias 公式只取校准值 | 改为 `=校准-REF` |
| Wafer Map 闪退 | 改回原版 `DrawMapSubWindow` 独立窗口 |
| Tab 5 透视表不可用 | xlsxwriter 无 `add_pivot_table`，取消 Tab |

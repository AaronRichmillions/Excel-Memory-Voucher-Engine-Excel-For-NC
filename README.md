# Excel Memory-Level Voucher Engine v3.0 
# Excel 内存级财务凭证引擎 v3.0

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## 📖 Introduction / 项目简介
This project features a sophisticated financial auditing engine built entirely within Excel's `LET` function. By establishing a "Virtual Memory Computing Layer," it automates financial voucher generation and anomaly detection without the need for physical helper columns.

本项目通过 Excel 的 `LET` 函数构建了一个硬核的“虚拟内存计算层”。它能在内存中完成财务分录生成与异常审计，彻底摆脱了传统的辅助列堆砌。

---

## 🚀 Core Features / 核心功能
- **Anchor Positioning (Single-Point Config)**: Define one column, and the engine auto-calculates all related column offsets.
  **单点锚定**：只需指定一个核心列标，引擎即可自动推导全行关联数据。
- **Dynamic Spill Logic**: Utilizes Excel's spill range technology to generate multi-line vouchers from single-row inputs.
  **动态溢出逻辑**：利用 Spill 机制，将单行原始数据在内存中“裂变”为标准财务分录。
- **Zero Artifacts**: No messy intermediate data; the output is clean and memory-efficient.
  **零碎屑**：计算过程全在内存完成，不会在工作表中留下任何中间辅助数据。

---

## 🛠️ Configuration / 使用配置
Simply modify the variables at the top of the formula:
只需修改公式顶部的变量即可适配你的表格：

```excel
  Row_Var, 45:100,         /* Global Row Range / 全局行号变量 */
  Anchor_Col, "T",         /* Match Anchor / 匹配锚点列标 */
⚠️ Important Note / 注意事项
Spill Reference: Since the output is a dynamic array, referencing specific cells requires the # symbol (e.g., =O46#). Direct references like =O75 may return 0 because those cells are "shadows" of the memory array.

溢出引用说明：由于输出结果是动态数组，引用结果时必须使用 # 符号（例如 =O46#）。直接引用如 =O75 可能会返回 0，因为在物理上那些单元格只是内存数组的“影子”。

⚖️ License / 授权协议
MIT License - Feel free to use and modify for your financial workflows. 本项目采用 MIT 协议 - 欢迎在财务流程中自由使用和修改。

# GMV File Auto Organizer

Build status: successfully built on 2026-06-01.

## Package Contents

- `自动GMV整理.exe`
  - Windows executable version.
  - No Python installation is required.
  - Launches without a CMD console window.

- `app.py`
  - Source code for the GMV organizer.

- `店铺-运营匹配表模版.xlsx`
  - Template for store-to-operator mapping.

- `config.json`
  - Stores the last used source data folder only.

- `启动程序.bat`
  - Optional Python-based launcher for development or fallback use.

- `requirements.txt`
  - Python dependency list for source-code execution.

## Core Logic

The program organizes GMV-related files by store and operator.

It reads store folders from the source data directory, matches each store to an operator through the store-operator mapping table, filters out empty files that only contain a header row, and outputs files under an operator-based directory structure.

Output structure:

```text
Target Output Folder/
  Operator Name/
    订单明细/
    广告明细/
    广告消耗明细/
    退款明细/
    推广明细/
```

Store matching is case-insensitive.

Promotion order data is split by store using column B, and only stores that exist in the store-operator mapping table are exported.

## Build Note

This package was successfully built as a standalone Windows executable on 2026-06-01.

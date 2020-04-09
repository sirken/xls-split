# XLS Split
Split an XLS into multiple files, retaining headers.

## Installation

1. Install `openpyxl`
    ```bash
    pip3 install openpyxl
    ```

2. Set execute permissions for script
    ```bash
    chmod +x xls-split
    ```

## Usage

1. Split target xls into new files containing x number or rows:
    ```bash
    ./xls-split "file.xls" 500
    ```

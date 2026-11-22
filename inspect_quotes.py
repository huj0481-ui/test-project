"""
第一步：查看 quotes/ 文件夹下所有供应商报价单的原始结构。
只读取和打印，不做任何数据清洗。
"""

import sys
from pathlib import Path

import pandas as pd

# 报价单所在文件夹：默认是脚本旁边的 quotes/，也可以在命令行传入其他路径
QUOTES_DIR = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).parent / "quotes"


def inspect_file(path: Path) -> None:
    """读取单个 xlsx 文件并打印其基本信息。"""
    print("=" * 60)
    print(f"文件名：{path.name}")

    try:
        # 只读第一个工作表；header=0 表示把第一行当作列名
        df = pd.read_excel(path, sheet_name=0, header=0)
    except Exception as e:  # 文件损坏、加密、格式不对等情况都在这里捕获
        print(f"[读取失败] {type(e).__name__}: {e}")
        return

    print(f"列名列表：{list(df.columns)}")
    print(f"总行数：{len(df)}")
    print("前 5 行数据：")
    print(df.head(5).to_string())


def main() -> None:
    if not QUOTES_DIR.is_dir():
        print(f"找不到文件夹：{QUOTES_DIR.resolve()}")
        return

    # 找出所有 xlsx，跳过 Excel 打开文件时产生的临时文件（以 ~$ 开头）
    files = sorted(p for p in QUOTES_DIR.glob("*.xlsx") if not p.name.startswith("~$"))
    if not files:
        print(f"{QUOTES_DIR.resolve()} 下没有 .xlsx 文件")
        return

    print(f"在 {QUOTES_DIR.resolve()} 找到 {len(files)} 个报价单\n")
    for path in files:
        inspect_file(path)


if __name__ == "__main__":
    # 让中文列名在终端里对齐得更好看一些
    pd.set_option("display.unicode.east_asian_width", True)
    main()

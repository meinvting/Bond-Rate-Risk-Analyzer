# -*- coding: utf-8 -*-
"""
验证脚本：读取鲍玉婷的 Excel 工作簿，校验 B9 单元格计算结果是否等于 965.35。
输出：验证通过 / 验证失败

用法：
    python verify_b9.py
    python verify_b9.py "D:/path/to/your.xlsx"   # 可传入自定义路径

说明：
    B9 为 =PV(...) 定价公式，结果应为 965.3489...（四舍五入 965.35）。
    若文件曾被非 Excel 程序（如 openpyxl）重写导致公式缓存值丢失，
    脚本会自动回退到读取 B2:B6 输入参数并按同一 PV 逻辑重算，保证仍可验证。
"""

import sys
import os
import glob

try:
    import openpyxl
except ImportError:
    print("验证失败：未安装 openpyxl，请先执行 pip install openpyxl")
    sys.exit(1)

# 默认路径（如文件位置变动，可在此修改或通过命令行参数传入）
DEFAULT_PATH = (
    r"C:/Users/DELL/OneDrive/桌面/"
    r"3202610086 金专3班 鲍玉婷 债券利率风险分析器/"
    r"3202610086 金专3班 鲍玉婷 债券利率风险分析器2.xlsx"
)

TARGET = 965.35          # 期望的教材例题 12-3 校验值
TOL_ROUND = 2            # 保留两位小数后比较（避免因浮点精度 965.3489... 误判）


def resolve_path(arg_path):
    """优先使用传入路径；否则尝试默认路径；再不行用通配符容错定位。"""
    if arg_path and os.path.exists(arg_path):
        return arg_path
    if os.path.exists(DEFAULT_PATH):
        return DEFAULT_PATH
    hits = glob.glob(r"C:/Users/DELL/OneDrive/桌面/*鲍玉婷*债券利率*/*.xlsx")
    return hits[0] if hits else None


def compute_pv(face, coup, years, freq, yld):
    """与 Excel PV 等价的债券定价：面值、票息率、期限(年)、年付息次数、到期收益率。"""
    r = yld / freq
    n = years * freq
    pmt = -(face * coup / freq)
    fv = -face
    return -(pmt * (1 - (1 + r) ** (-n)) / r + fv / ((1 + r) ** n))


def main():
    path = resolve_path(sys.argv[1] if len(sys.argv) > 1 else None)
    if not path:
        print("验证失败：未找到 Excel 文件，请检查路径")
        return

    # 1) 先尝试读取公式计算后的缓存值
    val = None
    src = "单元格缓存值"
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        val = wb.active["B9"].value
    except Exception as e:
        print(f"验证失败：无法打开文件（{e}）")
        return

    # 2) 若缓存缺失（被非 Excel 程序重写过），回退到按 B2:B6 参数重算
    if val is None or not isinstance(val, (int, float)):
        try:
            wbf = openpyxl.load_workbook(path)  # 公式模式
            wsf = wbf.active
            face = float(wsf["B2"].value)
            coup = float(wsf["B3"].value)
            years = float(wsf["B4"].value)
            freq = float(wsf["B5"].value)
            yld = float(wsf["B6"].value)
            val = compute_pv(face, coup, years, freq, yld)
            src = "按 B2:B6 输入参数重算（公式缓存缺失）"
        except Exception as e:
            print(f"验证失败：B9 无计算结果且无法读取输入参数重算（{e}）")
            return

    try:
        num = float(val)
    except (TypeError, ValueError):
        print(f"验证失败：B9 内容非数值（{val!r}）")
        return

    # 浮点比较：四舍五入到 2 位小数后等于 965.35 即视为通过
    if round(num, TOL_ROUND) == TARGET:
        print(f"验证通过（依据：{src}，B9≈{round(num, 4)}）")
    else:
        print(f"验证失败：B9 = {num}，不等于 {TARGET}（依据：{src}）")


if __name__ == "__main__":
    main()

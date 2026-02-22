# -*- coding: utf-8 -*-
"""Dump programs grouped for minute chart pages."""
from data_reader import read_excel_data
import sys
sys.path.insert(0, '.')
from generate_report_config_driven import _group_programs

data = read_excel_data("input_data/20260215CCTV-17日数据（简版）.xlsx")
groups = _group_programs(data.programs, max_cols=20)
print(f"Total groups: {len(groups)}")
for i, g in enumerate(groups):
    print(f"  [{i:2d}] {g['start']:4d}-{g['end']:4d} dur={g['duration']:4d}  {g['name'][:50]}")

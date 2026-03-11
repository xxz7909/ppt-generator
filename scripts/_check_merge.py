"""检查节目合并结果"""
from data_reader import read_excel_data
from generate_report_config_driven import _group_programs
import sys

data = read_excel_data(r'input_data\20260310CCTV-17日数据.xlsx')

print("--- max_cols=35 (修改后) ---")
groups = _group_programs(data.programs, max_cols=35)
print(f"Grouped to {len(groups)} items:")
merged_count = 0
for g in groups:
    name = g['name']
    dur = g['duration']
    start = g['start']
    end = g['end']
    has_plus = '+' in name
    if has_plus:
        merged_count += 1
    flag = ' *** MERGED ***' if has_plus else ''
    print(f"  [{start}-{end}] ({dur}min) {name}{flag}")
print(f"\nMerged items: {merged_count}")
print()

print("--- max_cols=20 (修改前) ---")
groups_old = _group_programs(data.programs, max_cols=20)
print(f"Grouped to {len(groups_old)} items:")
merged_old = sum(1 for g in groups_old if '+' in g['name'])
print(f"Merged items: {merged_old}")

"""
Test script for Excel Parser
"""
from excel_parser import parse_excel_bookings

files = [
    r'C:\Users\HP\Desktop\30389 - GC - 15-09- 26 - QAI Bookings.xlsx',
    r'C:\Users\HP\Desktop\37931 - Pneumatic Checking - 31-08-26 - LBR and Other Bookings.xlsx',
    r'C:\Users\HP\Desktop\30528 - 31-08-26 - LBR and Other Bookings.xlsx',
    r'C:\Users\HP\Desktop\30577 - IA - 31-08-26 - LBR and PST Bookings.xlsx'
]

for f in files:
    res = parse_excel_bookings(f)
    print("=" * 60)
    print("File:", res['file_name'])
    print("Loco:", res['loco_number'], "| Date:", res['date'], "| Schedule:", res['schedule'])
    print("Total Bookings:", res['bookings_count'])
    for b in res['bookings'][:3]:
        print(f"   S.No {b['sno']}: [{b['raw_section']} -> {b['sections']}] {b['description'][:50]}...")
    if res['bookings_count'] > 3:
        last_b = res['bookings'][-1]
        print(f"   ... S.No {last_b['sno']}: [{last_b['raw_section']} -> {last_b['sections']}] {last_b['description'][:50]}...")

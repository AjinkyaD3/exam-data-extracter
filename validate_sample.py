import json
import random

# Load the Sem 3 output
with open('output/AI_DS_Sem3_Winter_2025.json', 'r') as f:
    data = json.load(f)

students = data['students']

# Find 3 PASS students
pass_students = [s for s in students if s['result'] == 'PASS']
sample_indices = [0, len(pass_students)//2, -1]

print("=== LEVEL 1 & 2 VALIDATION CHECK (PASS STUDENTS) ===")
for idx in sample_indices:
    if idx < len(pass_students):
        s = pass_students[idx]
        print(f"\nStudent: {s['seat_no']} - {s['name']}")
        print(f"Result: {s['result']}")
        print(f"SGPA: {s['sgpa']}")
        print(f"Total Marks: {s['total_marks']}")
        print("-" * 30)
        # Print subject totals
        print("Subject Totals:")
        for key, value in s.items():
            if key.endswith('_total') and value is not None:
                subject_code = key.replace('_total', '')
                print(f"  {subject_code}: {value}")

import pandas as pd
import random

def randomize_groups():
    file_path = input("Enter Excel filename (example: students.xlsx): ")
    df = pd.read_excel(file_path)

    students = df['Name'].tolist()

    total_students = int(input("How many students are enrolled? "))
    total_groups = int(input("How many groups? "))

    if len(students) != total_students:
        print("\nNotice: The Excel file has", len(students), "students.")
        print("But you entered:", total_students)
        print("Using the number of students found in the file.\n")
        total_students = len(students)
        print(total_students)

    random.shuffle(students)

    groups = {f"Group {i+1}": [] for i in range(total_groups)}

    group_index = 0
    for student in students:
        group_name = f"Group {group_index + 1}"
        groups[group_name].append(student)

        group_index = (group_index + 1) % total_groups

    print("\n=== RANDOMIZED GROUPS ===")
    for group, members in groups.items():
        print(f"\n{group} ({len(members)} students):")
        for m in members:
            print("  -", m)

    save_choice = input("\nDo you want to export results to Excel? (y/n): ")
    if save_choice.lower() == 'y':
        output = []
        for group, members in groups.items():
            for m in members:
                output.append({'Group': group, 'Student': m})
        result_df = pd.DataFrame(output)
        result_df.to_excel("group_results.xlsx", index=False)
        print("Saved as group_results.xlsx")

randomize_groups()
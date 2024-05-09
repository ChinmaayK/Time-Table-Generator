from Generate import generate

# generate dummy data like above from classes 10 to 1 A,B,C
nine_ten = ["Maths", "Phy", "Chem", "Bio", "His", "Geo", "Eng1", "Eng2", "Hindi", "Comp"]
data = {
    "10A": [nine_ten,
            ["Lisha", "Divyam", "Seema", "Vaishali", "Presentina", "Zakira", "Aparna", "Vineeta", "Sunanda", "SI"],
                [7, 4, 4, 4, 3, 3, 3, 4, 4, 4], "Lisha"],
    "10B": [nine_ten, ["Lisha", "SPhy", "Seema", "Usha", "Aparna", "Zakira", "Aparna", "Vineeta", "Hirdesh", "Merry"],
            [7, 4, 4, 4, 3, 3, 3, 4, 4, 4], "Zakira"],
    "10C": [nine_ten,
            ["Pratiksha", "SPhy", "Seema", "Usha", "Aparna", "Zakira", "Aparna", "Vineeta", "Hirdesh", "Merry"],
            [7, 4, 4, 4, 3, 3, 3, 4, 4, 4], "Usha"],
    "9A": [nine_ten,
           ["Pratiksha", "SPhy", "Seema", "Vaishali", "Aparna", "Zakira", "Aparna", "Vineeta", "Hirdesh", "Merry"],
           [7, 4, 4, 4, 3, 3, 3, 4, 4, 4], "Merry"],
    "9B": [nine_ten, ["Pratiksha", "SPhy", "Seema", "Usha", "Presentina", "Zakira", "Vineeta", "Arti", "Hirdesh", "SI"],
           [7, 4, 4, 4, 3, 3, 3, 4, 4, 4], "Pratiksha"],
    "9C": [nine_ten,
           ["Pratiksha", "SPhy", "Seema", "Usha", "Aparna", "Zakira", "Vineta", "Vineeta", "Hirdesh", "Merry"],
           [7, 4, 4, 4, 3, 3, 3, 4, 4, 4], "Vineeta"],
}


def chk_overloaded_teachers(data, overload_value):
    teacher_frequency = {}

    for class_data in data.values():
        teachers = class_data[1]
        frequencies = class_data[2]

        for teacher, frequency in zip(teachers, frequencies):
            teacher_frequency[teacher] = teacher_frequency.get(teacher, 0) + frequency

    overloaded_teachers = {teacher: frequency for teacher, frequency in teacher_frequency.items() if
                           frequency > overload_value}
    overloaded_students = [class_name for class_name, class_data in data.items() if sum(class_data[2]) > overload_value]

    return overloaded_teachers, overloaded_students


def gen_matrix(rows, columns, days, data):
    # [days*[column[*rows]]]
    matrix = [{} for _ in range(days)]

    for m in matrix:
        # matrix[m] keys = keys of data and values = [0,0,0]
        for key, value in data.items():
            m[key] = [0 for _ in range(rows)]

    return matrix


def gen_teacher_data(data):
    teacher_data = {}
    for class_name, class_data, in data.items():
        teachers = class_data[1]
        for idx, teacher in enumerate(teachers):
            teacher_data[teacher] = teacher_data.get(teacher, []) + [[class_name, data[class_name][0][idx], data[class_name][2][idx]]]

    return teacher_data


def gen_extra_classes(data, days):
    extra_classes = {}
    for idx, (key, values) in enumerate(data.items()):
        freq_list = []
        for index, freq in enumerate(values[2]):
            if freq > days and freq <= 2 * (days):
                freq_list.append([values[0][index], freq-days, 1])
            elif freq > days and freq > 2 * (days):
                freq_list.append([values[0][index], freq-days, int(freq/days)+1])
        if freq_list != []:
            extra_classes[key] = freq_list
    return extra_classes


def generate_teachers_tt(matrix):
    teacher_timetables = {}

    # Loop through each day's data (each dictionary in the list)
    for day_idx, day_data in enumerate(matrix):
        # Loop through each class/division
        for class_name, periods in day_data.items():
            # Loop through each period
            for period_idx, period_info in enumerate(periods):
                teacher_name, subject = period_info

                # Check if the teacher already has a timetable
                if teacher_name not in teacher_timetables:
                    # Initialize a new timetable with 5 empty days
                    teacher_timetables[teacher_name] = {f"Day {i + 1}": [0] * 8 for i in range(5)}

                # Update the teacher's timetable for this day and period
                teacher_timetables[teacher_name][f"Day {day_idx + 1}"][period_idx] = (subject, class_name)

    return teacher_timetables
    # # Now, teacher_timetables contains the schedule for each teacher
    # for teacher_name, timetable in teacher_timetables.items():
    #     print(f"Timetable for {teacher_name}:")
    #     for day, periods in timetable.items():
    #         print(f"{day}: {periods}")
    #     print()  # New line for separation


def generate_class_tt(matrix):
    timetables = {}

    # Check if matrix contains only dictionaries
    for day in matrix:
        if not isinstance(day, dict):
            print(f"Error: Expected dictionary, but got {type(day)}")
            continue  # Skip if not a dictionary

        # Find all unique class names
        for class_name in day.keys():
            if class_name not in timetables:
                timetables[class_name] = {}  # Initialize the timetable for each class

    # Organize the schedules for each class
    for day_idx, day in enumerate(matrix):
        if not isinstance(day, dict):
            continue  # Skip if not a dictionary

        day_key = f"Day {day_idx + 1}"

        for class_name, schedule in day.items():
            if day_key not in timetables[class_name]:
                timetables[class_name][day_key] = []  # Initialize an empty list

            timetables[class_name][day_key].extend(schedule)  # Add schedules for the day

    return timetables


def run(data, total_periods, total_days):
    overloaded_teachers, overloaded_stu = chk_overloaded_teachers(data, total_periods * total_days)
    if not overloaded_stu and not overloaded_teachers:
        matrix = gen_matrix(total_periods, len(data), total_days, data)
        teacher_data = gen_teacher_data(data)
        extra_classes = gen_extra_classes(data, total_days)
        # print(extra_classes)
        try:
            matrix, teachers_data, data, = generate(matrix, teacher_data, data, extra_classes)  # Ensure generate returns the correct tuple
            teacher_timetables = generate_teachers_tt(matrix)
            class_timetables = generate_class_tt(matrix)

            return matrix, teacher_timetables, class_timetables
        except Exception as e:
            print("An error occurred:", str(e))

    else:
        print("Error due to ")
        print(overloaded_stu)
        print(overloaded_teachers)


if __name__ == '__main__':
    run(data, 8, 5)

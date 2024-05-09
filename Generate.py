def get_available_teachers(class_name, teachers_data, class_teacher, matrix, index, extra_classes):
    available_teachers = []
    final_at = []
    ct_freq = False
    ct_subj = False

    for teacher, teacher_data in teachers_data.items():
        for cls in teacher_data:
            if cls[0] == class_name and cls[2] > 0:
                if teacher == class_teacher:
                    ct_freq = cls[2]
                    ct_subj = cls[1]
                available_teachers.append([teacher, cls[2], cls[1]])

    available_teachers.sort(key=lambda x: x[1], reverse=True)
    if ct_freq is not False:
        available_teachers.remove([class_teacher, ct_freq, ct_subj])
        available_teachers.insert(0, [class_teacher, ct_freq, ct_subj])

    # print(available_teachers)
    extra_classes2 = extra_classes
    for n in available_teachers:
        extra_classes = chk_available_teachers(matrix, n[0], index, class_name, extra_classes, n[2])
        if extra_classes is not False:
            final_at.append(n)
            extra_classes2 = extra_classes
        else:
            extra_classes = extra_classes2

    # print(final_at)
    return final_at, extra_classes


def chk_available_teachers(matrix, teacher, idx, class_name, extra_classes, subject):
    for key, values, in matrix.items():
        if values[idx] != 0:
            if values[idx][0] == teacher:
                return False
    counter = 0
    for index, values, in enumerate(matrix[class_name]):
        if index < idx:
            if matrix[class_name][idx-1] == [teacher, subject]:
                return False
            if values != 0:
                if values[0] == teacher:
                    counter += 1
    if counter > 0:
        for subj in extra_classes[class_name]:
            if subj[0] == teacher and subj[1] > 0 and subj[2] > counter-1:
                subj[1] -= 1
                return extra_classes
            elif subj[1] < 0 or subj[2] <= counter-1:
                return False
    return extra_classes


def get_empty_cell(matrix):
    for idx_m, m in enumerate(matrix):
        for key, values in m.items():
            for idx, cls in enumerate(values):
                if cls == 0:
                    return idx_m, key, idx
    return None, None, None


def generate(matrix, teachers_data, data, extra_classes):
    day, class_name, pos = get_empty_cell(matrix)
    if day is None:
        return matrix, teachers_data, data

    available_teachers, extra_classes = get_available_teachers(class_name, teachers_data, data[class_name][3], matrix[day], pos, extra_classes)
    if available_teachers != []:
        for at in available_teachers:
            matrix[day][class_name][pos] = [at[0], at[2]]
            for teacher in teachers_data[at[0]]:
                if teacher[0] == class_name and teacher[1] == at[2]:
                    teacher[2] -= 1
            if generate(matrix, teachers_data, data, extra_classes) is not False:
                return matrix, teachers_data, data

            matrix[day][class_name][pos] = 0
            for teacher in teachers_data[at[0]]:
                if teacher[0] == class_name and teacher[1] == at[2]:
                    teacher[2] += 1
    return False


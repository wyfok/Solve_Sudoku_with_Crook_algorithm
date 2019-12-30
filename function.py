import copy
import pandas as pd


def read_question_from_excel(_excel_name, _sheet_name):
    df = pd.read_excel(_excel_name, sheet_name=_sheet_name, header=None)
    return df.values.tolist()


def check_row(possible_value_, solution_):
    """
    :param possible_value_: the dict of storing all possible numbers of each cell
    :param solution_: the list of existing solution
    Based on Sudoku rule, check for each row and see if there is an unique possible number across a row.
    If so, update the possible_value_ and solution.
    """
    for i in range(1, 10):
        exist = solution_[i - 1]

        for j in range(1, 10):
            possible_value_[i, j] = [x for x in possible_value_[i, j] if x not in exist]

        possible_element = [x for y in [value for key, value in possible_value_.items()
                                        if key[0] == i and len(value) > 0] for x in y]
        unique = [x for x in possible_element if possible_element.count(x) == 1]
        if len(unique) > 0:
            for x in unique:
                for key, value in {key: value for key, value in possible_value_.items() if
                                   key[0] == i and len(value) > 0}.items():
                    if x in value:
                        solution_[key[0] - 1][key[1] - 1] = x
                        possible_value_[key] = []
    return 0


def check_column(possible_value_, solution_):
    """
    :param possible_value_: the dict of storing all possible numbers of each cell
    :param solution_: the list of existing solution
    Based on Sudoku rule, check for each column and see if there is an unique possible number across a column.
    If so, update the possible_value_ and solution_.
    """
    for j in range(1, 10):
        exist = [x[j - 1] for x in solution_]
        for i in range(1, 10):
            possible_value_[i, j] = [x for x in possible_value_[i, j] if x not in exist]

        possible_element = [x for y in [value for key, value in possible_value_.items()
                                        if key[1] == j and len(value) > 0] for x in y]
        unique = [x for x in possible_element if possible_element.count(x) == 1]
        if len(unique) > 0:
            for x in unique:
                for key, value in {key: value for key, value in possible_value_.items() if
                                   key[1] == j and len(value) > 0}.items():
                    if x in value:
                        solution_[key[0] - 1][key[1] - 1] = x
                        possible_value_[key] = []
    return 0


def box_range(number):
    """
    :param number: input the row or column number
    :return: a list of row or column number within the same box
    """
    if number in (1, 2, 3):
        return [1, 2, 3]
    elif number in (4, 5, 6):
        return [4, 5, 6]
    elif number in (7, 8, 9):
        return [7, 8, 9]


def check_box(possible_value_, solution_):
    """
    :param possible_value_: the dict of storing all possible numbers of each cell
    :param solution_: the list of existing solution
    Based on Sudoku, check for each box and see if there is an unique possible number within a box.
    If so, update the possible_value_ and solution_.
    """
    for i in [1, 4, 7]:
        for j in [1, 4, 7]:
            exist = set(
                [solution_[i_range - 1][j_range - 1] for j_range in range(j, j + 3) for i_range in range(i, i + 3)])
            for k in box_range(i):
                for l in box_range(j):
                    possible_value_[k, l] = [b for b in possible_value_[k, l] if b not in exist]

            possible_element = [x for b in [value for key, value in possible_value_.items()
                                            if key[0] in box_range(i) and key[1] in box_range(j) and len(value) > 0]
                                for x in b]
            unique = [x for x in possible_element if possible_element.count(x) == 1]
            if len(unique) > 0:
                for k in unique:
                    for key, value in {key: value for key, value in possible_value_.items()
                                       if key[0] in box_range(i) and key[1] in box_range(j) and len(value) > 0}.items():
                        if k in value:
                            solution_[key[0] - 1][key[1] - 1] = k
                            possible_value_[key] = []
    return 0


def check_unique_possible_value(possible_value_, solution_):
    """
    :param possible_value_: the dict of storing all possible numbers of each cell
    :param solution_: the list of existing solution
    For each cell, if there is only one possible number, update solution_ and remove from possible_value_
    """
    for key, value in possible_value_.items():
        if len(value) == 1:
            solution_[key[0] - 1][key[1] - 1] = value[0]
            possible_value_[key] = []
    return 0


def loop_basic_rule(possible_value_, solution_):
    """
    :param possible_value_: the dict of storing all possible numbers of each cell
    :param solution_: the list of existing solution
    Run all basic rules until there is no update on solution_
    """
    while True:
        solution_old = copy.deepcopy(solution_)
        check_row(possible_value_, solution_)
        check_column(possible_value_, solution_)
        check_box(possible_value_, solution_)
        check_unique_possible_value(possible_value_, solution_)
        if solution_ == solution_old:
            break


def algorithm(possible_list_, possible_value_, solution_):
    """
    :param possible_list_: a markup of a cell
    :param possible_value_: the dict of storing all possible numbers of each cell
    :param solution_: the list of existing solution
    The function applies Crook's algorithm and eliminate impossible numbers.
    If there is any update, re-run all basic rules
    """
    try:
        min_num_possible = min((len(v)) for _, v in possible_list_.items())
        max_num_possible = max((len(v)) for _, v in possible_list_.items())
    except ValueError:
        return 0
    for i in reversed(range(min_num_possible, max_num_possible + 1)):
        for key, value in {key: value for key, value in possible_list_.items() if len(value) == i}.items():
            n_subset = 0
            key_match = set()
            for key_1, value_1 in possible_list_.items():
                if len(value) < len(value_1):
                    continue
                else:
                    if set(value_1).issubset(set(value)):
                        key_match.add(key_1)
                        n_subset += 1
                if n_subset == len(value):
                    for key_2, value_2 in {key: value for key, value in possible_list_.items() if
                                           key not in key_match}.items():
                        possible_value_[key_2] = [x for x in value_2 if x not in value]
                        loop_basic_rule(possible_value_, solution_)
    return 0


def algorithm_row(possible_value_, solution_):
    """
    :param possible_value_: the dict of storing all possible numbers of each cell
    :param solution_: the list of existing solution
    Apply Crook's algorithm on each row
    """
    for i in range(1, 10):
        possible_list = {key: value for key, value in possible_value_.items() if key[0] == i and len(value) > 0}
        algorithm(possible_list, possible_value_, solution_)
    return 0


def algorithm_column(possible_value_, solution_):
    """
    :param possible_value_: the dict of storing all possible numbers of each cell
    :param solution_: the list of existing solution
    Apply Crook's algorithm on each column
    """
    for j in range(1, 10):
        possible_list = {key: value for key, value in possible_value_.items() if key[1] == j and len(value) > 0}
        algorithm(possible_list, possible_value_, solution_)
    return 0


def algorithm_box(possible_value_, solution_):
    """
    :param possible_value_: the dict of storing all possible numbers of each cell
    :param solution_: the list of existing solution
    Apply Crook's algorithm on each box
    """
    for i in [1, 4, 7]:
        for j in [1, 4, 7]:
            possible_list = {key: value for key, value in possible_value_.items() if
                             key[0] in box_range(i) and key[1] in box_range(j) and len(value) > 0}
            algorithm(possible_list, possible_value_, solution_)
    return 0


def loop_algorithm(possible_value_, solution_):
    """
    :param possible_value_: the dict of storing all possible numbers of each cell
    :param solution_: the list of existing solution
    Apply Crook's algorithm until there is no update on solution_
    """
    while True:
        solution_old = copy.deepcopy(solution_)
        algorithm_row(possible_value_, solution_)
        algorithm_column(possible_value_, solution_)
        algorithm_box(possible_value_, solution_)
        if solution_ == solution_old:
            break
    return 0


def check_box_eliminate_others(possible_value_):
    """
    :param possible_value_: the dict of storing all possible numbers of each cell
    By considering the possible numbers within a box, check if there is a possible number only in one row/column.
    If so, then in the same row/column outside the box, this possible number will be eliminated from all markup.
    """
    for i in [1, 4, 7]:
        for j in [1, 4, 7]:
            possible_element = set([x for b in
                                    [value for key, value in possible_value_.items()
                                     if key[0] in box_range(i) and key[1] in box_range(j) and len(value) > 0]
                                    for x in b])

            for x in possible_element:
                available_cell = [key for key, value in possible_value_.items()
                                  if x in value if key[0] in box_range(i) and key[1] in box_range(j)]
                if len(set([x[0] for x in available_cell])) == 1:
                    for key in [key for key, value in possible_value_.items() if
                                key[0] == available_cell[0][0] and key not in available_cell]:
                        possible_value_[key] = [y for y in possible_value_[key] if y != x]
                if len(set([x[1] for x in available_cell])) == 1:
                    for key in [key for key, value in possible_value_.items() if
                                key[1] == available_cell[0][1] and key not in available_cell]:
                        possible_value_[key] = [y for y in possible_value_[key] if y != x]
    return 0

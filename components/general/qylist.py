"""

By Ziqiu Li
Created at 2023/3/8 16:47
"""


def add_one_to_list(src_list: list, one: object):
    src_list.append(one)


def reverse_list(src_list: list):
    src_list.reverse()


def concat_list(src_list: list, r_list: list):
    return src_list + r_list


def sort_list(src_list: list, reverse: str = "升序"):
    if reverse == "降序":
        src_list.sort(reverse=True)
    else:
        src_list.sort()


def transpose_list(src_list: list):
    return list(map(list, zip(*src_list)))


def update_one(src_list: list, idx: int, value):
    # idx = idx - 1 if idx > 0 else idx
    src_list[idx] = value


def get_data_from_index(src_list: list, idx: int):
    return src_list[idx]


def get_data_index(src_list: list, data):
    return src_list.index(data)


def duplicate_list(src_list: list):
    return list(set(src_list)).sort(key=src_list.index)


def filter_list(src_list: list, filter_data: list):
    return [i for i in src_list if i not in filter_data]


def get_same_data(src_list1: list, src_list2: list):
    return list(set(src_list1) & set(src_list2))


def get_max_data(src_list: list):
    return max(src_list)


def get_min_data(src_list: list):
    return min(src_list)


def delete_none_data(src_list: list):
    none_list = ["", None]
    return filter_list(src_list, none_list)


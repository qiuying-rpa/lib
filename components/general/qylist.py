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


def sort_list(src_list: list, reverse: str):
    if reverse == "descending":
        src_list.sort(reverse=True)
    else:
        src_list.sort()

"""

By Ziqiu Li
Created at 2023/3/15 17:42
"""


def add_key_value(src_dict: dict, key, value):
    src_dict.update({
        key: value
    })


def get_value(src_dict: dict, key, raise_fault: bool = True, default_value=""):
    if raise_fault:
        return src_dict[key]
    else:
        return src_dict.get(key, default_value)


def get_key_list(src_dict: dict):
    return src_dict.keys()


def get_value_list(src_dict: dict):
    return src_dict.values()


def delete_dict_by_key(src_dict: dict, key):
    del src_dict[key]




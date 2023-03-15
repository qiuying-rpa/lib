"""

By Ziqiu Li
Created at 2023/3/8 16:06
"""


def split_string_to_list(text: str, sep: str) -> list:
    """

    :param text:
    :param sep:
    :return:
    """
    return text.split(sep=sep)


def join_list_to_string(str_list: list, join_str: str) -> str:
    """

    :param str_list:
    :param join_str:
    :return:
    """
    return join_str.join([str(i) for i in str_list])


def replace_string(text: str, old_str, new_str) -> str:
    """

    :param text:
    :param old_str:
    :param new_str:
    :return:
    """
    return text.replace(old_str, new_str)


def get_string_length(text: str) -> int:
    """

    :param text:
    :return:
    """
    return len(text)


def slice_string(text: str, start_index: int, end_index: int) -> str:
    """

    :param text:
    :param start_index:
    :param end_index:
    :return:
    """
    # start_index = start_index - 1 if start_index > 0 else start_index
    # end_index = end_index - 1 if end_index > 0 else end_index
    return text[start_index, end_index]


def trans_upper_or_lower(text: str, trans_type: str) -> str:
    """

    :param text:
    :param trans_type: ["upper", "lower"]
    :return:
    """
    if trans_type == "upper":
        return text.upper()
    else:
        return text.lower()


def concat_text(text: str, r_text: str) -> str:
    """

    :param text:
    :param r_text:
    :return:
    """
    return text + r_text

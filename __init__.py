"""

By Allen Tao
Created at 2023/01/18 13:47
"""


def get_components_info():
    """
    按照约定，从注释中获取组件信息
    """
    import importlib
    module = importlib.import_module('components.km')
    print(module)
    func = getattr(module, 'key_press')
    print(func.__doc__)


if __name__ == '__main__':
    get_components_info()

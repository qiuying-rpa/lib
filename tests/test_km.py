"""
Some tests.

By Allen Tao
Created at 2023/3/17 15:56
"""
from qiuying.components.km import km


def test_km():
    km.keys_press(['win', 'r'], 1)
    km.type_string('cmd', .5)
    km.key_press('enter', 1)
    km.type_string('Don\'t answer!')

from qiuying.components.km import km


def test_km():
    km.keys_press(['win', 'r'], 1)
    km.type_string('cmd', .5)
    km.key_press('enter', 1)
    km.type_string('Don\'t answer!')

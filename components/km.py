"""
Keyboard and mouse related.

By Allen Tao
Created at 2022/09/26 14:00
"""
import time
from pynput import keyboard, mouse
from pynput.keyboard import Key
from models.error import InvalidMouseButtonException

_mouse_controller = mouse.Controller()
_keyboard_controller = keyboard.Controller()


# -- keyboard


_key_map = {
    "command": Key.cmd,
    "cmd": Key.cmd,
    "win": Key.cmd,
    "enter": Key.enter,
    "esc": Key.esc,
    "tab": Key.tab,
    "caps_lock": Key.caps_lock,
    "shift": Key.shift,
    "ctrl": Key.ctrl,
    "alt": Key.alt,
    "space": Key.space,
    "up": Key.up,
    "down": Key.down,
    "left": Key.left,
    "right": Key.right,
    "backspace": Key.backspace,
    "end": Key.end,
    "page_down": Key.page_down,
    "page_up": Key.page_up,
    "home": Key.home,
    "menu": Key.menu,
    "insert": Key.insert,
    "delete": Key.delete,
    "num_lock": Key.num_lock,
    "print_screen": Key.print_screen,
    "scroll_lock": Key.scroll_lock,
    "pause": Key.pause,
    "ctrl_l": Key.ctrl_l,
    "ctrl_r": Key.ctrl_r,
    "cmd_l": Key.cmd_l,
    "cmd_r": Key.cmd_r,
    "shift_l": Key.shift_l,
    "shift_r": Key.shift_r,
    "alt_l": Key.alt_l,
    "alt_r": Key.alt_r,
    "alt_gr": Key.alt_gr,
    "f1": Key.f1,
    "f2": Key.f2,
    "f3": Key.f3,
    "f4": Key.f4,
    "f5": Key.f5,
    "f6": Key.f6,
    "f7": Key.f7,
    "f8": Key.f8,
    "f9": Key.f9,
    "f10": Key.f10,
    "f11": Key.f11,
    "f12": Key.f12
}


def key_press(key: str, delay_after: float = 0):
    """
    Press down and then release a key.

    :param key: the key
    :param delay_after: delay after pressing
    """
    key_down(key)
    key_up(key)
    time.sleep(delay_after)


def key_down(key: str, delay_after: float = 0):
    """Press down a key.

    :param key: the key
    :param delay_after: delay after pressing
    """
    _keyboard_controller.press(_key_map.get(key) or key)
    time.sleep(delay_after)


def key_up(key: str, delay_after: float = 0):
    """Release a key.

    :param key: the key
    :param delay_after: delay after releasing
    """
    _keyboard_controller.release(_key_map.get(key) or key)
    time.sleep(delay_after)


def type_string(string: str, delay_after: float = 0):
    """Type a string.

    :param string: the string
    :param delay_after: delay after typing
    """
    _keyboard_controller.type(string)
    time.sleep(delay_after)


def keys_press(keys: list, delay_after: float = 0):
    """Press and then release series of keys.

    :param keys: the keys, e.g. ["ctrl", "a"]
    :param delay_after: delay after releasing keys
    """
    # press each
    for key in keys:
        key_down(key)

    # release each
    for key in keys:
        key_up(key)

    time.sleep(delay_after)


# -- mouse


def get_mouse_position() -> tuple:
    """Get current position of mouse.

    :return: in a form of (x, y)
    """
    return _mouse_controller.position


def mouse_move(dx: int = 0, dy: int = 0, delay_after: float = 0,
               start_position: tuple[int, int] = (0, 0)):
    """Move mouse a certain distance from a certain position.
    As a default, `start_position` will be origin (left-top).

    :param dy: delta in y-axis to start x
    :param dx: delta in x-axis to start y
    :param delay_after: delay after moving
    :param start_position: the certain position. default is origin,
    """
    if start_position is not None:
        _mouse_controller.position = start_position
    _mouse_controller.move(dx, dy)
    time.sleep(delay_after)


def mouse_click(dx: int = 0, dy: int = 0, delay_after: float = 0, start_position: tuple[int, int] = (0, 0),
                button_type: str = 'left', double_click: bool = False):
    """Click mouse button at a certain position.

    :param dy: delta in y-axis to start x
    :param dx: delta in x-axis to start y
    :param delay_after: delay after clicking
    :param start_position: the position
    :param button_type: one of 'left', 'right', 'middle'
    :param double_click:
    :raises InvalidMouseButtonException:
    """
    if hasattr(mouse.Button, button_type):
        mouse_move(dx, dy, delay_after, start_position)
        _mouse_controller.click(getattr(mouse.Button, button_type), 2 if double_click else 1)
    else:
        raise InvalidMouseButtonException(button_type)


def mouse_scroll(dy: int = 0, dx: int = 0, delay_after: float = 0):
    """Mouse scroll.

    :param dy: delta in y-axis  positive means up, negative means down
    :param dx: delta in x-axis
    :param delay_after: delay after scrolling
    :return:
    """
    _mouse_controller.scroll(dx, dy)
    time.sleep(delay_after)

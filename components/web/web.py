"""

By Ziqiu Li
Created at 2023/2/15 14:23
"""

import time
from typing import Union
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException, \
    NoAlertPresentException, NoSuchWindowException


def _load_browser_default_option(download_folder=""):
    options = webdriver.ChromeOptions()
    perfs = {
        # 取消保存密码提示框
        "credentials_enable_service": False,
        "profile.password_manager_enabled": False,
        # 不屏蔽下载弹窗
        "profile.default_content_settings.popups": 1,
        # 设置默认下载路径
    }
    if download_folder:
        perfs.update({
            "download.default_directory": download_folder
        })
    options.add_experimental_option("prefs", perfs)
    options.add_experimental_option('detach', True)
    options.add_argument("disable-infobars")
    options.add_argument("no-sandbox")
    options.add_experimental_option("excludeSwitches", ['enable-automation'])
    return options


def _load_driver_default_setting(web_driver: WebDriver):
    # 页面最大加载时间60s  最大化窗口  隐式等待 30s
    web_driver.set_page_load_timeout(60)
    web_driver.maximize_window()
    web_driver.implicitly_wait(30)


def open_browser(download_folder: str = "") -> WebDriver:
    options = _load_browser_default_option(download_folder)
    web_driver = webdriver.Chrome(options=options)
    _load_driver_default_setting(web_driver)
    return web_driver


def open_url(web_driver: WebDriver, url: str) -> None:
    if "https://" in url or "http://" in url:
        web_driver.get(url)
    else:
        web_driver.get("http://" + url)


def element_is_exist(web_driver: WebDriver, locator: str, timeout: float = 10):
    try:
        WebDriverWait(web_driver, timeout).until(EC.presence_of_all_elements_located((By.XPATH, locator)))
        return True
    except TimeoutException as e:
        print(e)
        return False


def wait_until_element(web_driver: WebDriver, locator: str, wait_type: str, timeout: float = 10):
    begin_time = time.time()
    if wait_type == "visible":
        element_exist = element_is_exist(web_driver, locator, timeout)
        if element_exist:
            element = web_driver.find_element(By.XPATH, locator)
            current_time = time.time()
            while current_time - begin_time > timeout:
                if element.is_displayed():
                    return True
                time.sleep(1)
                current_time = time.time()
        else:
            return False
    else:
        element_exist = element_is_exist(web_driver, locator, timeout=2)
        if element_exist:
            element = web_driver.find_element(By.XPATH, locator)
            current_time = time.time()
            while current_time - begin_time > time.time():
                if not element.is_displayed():
                    return True
                time.sleep(1)
                current_time = time.time()
        else:
            return True
    return False


def scroll_to_visible(web_driver: WebDriver, locator: str, timeout: float = 10):
    element_exist = element_is_exist(web_driver, locator, timeout=timeout)
    if element_exist:
        try:
            ele = web_driver.find_element(By.XPATH, locator)
            web_driver.execute_script("arguments[0].scrollIntoView();", ele)
        except Exception as e:
            print("滚动至元素可见失败，失败原因：" + str(e))
    else:
        raise TimeoutException(f"查找元素：{locator}超时.")


def click_element(web_driver: WebDriver, locator: str, click_type: str = "单击", timeout: float = 10):
    element_exist = element_is_exist(web_driver, locator, timeout=timeout)
    if element_exist:
        ele = web_driver.find_element(By.XPATH, locator)
        if click_type == "双击":
            ActionChains(web_driver).double_click(ele).perform()
        elif click_type == "右击":
            ActionChains(web_driver).context_click(ele).perform()
        else:
            try:
                web_driver.execute_script("arguments[0].focus();", ele)
                ele.click()
            except StaleElementReferenceException as e:
                time.sleep(1)
                ele = web_driver.find_element((By.XPATH, locator))
                web_driver.execute_script("arguments[0].click();", ele)
    else:
        raise TimeoutException(f"查找元素：{locator}超时.")


def clear_element(web_driver: WebDriver, locator: str, timeout: float = 10):
    element_exist = element_is_exist(web_driver, locator, timeout=timeout)
    if element_exist:
        ele = web_driver.find_element(By.XPATH, locator)
        ele.clear()
    else:
        raise TimeoutException(f"查找元素：{locator}超时.")


def input_text(web_driver: WebDriver, locator: str, text: str, clear_flag: bool=True, timeout: float = 10):
    element_exist = element_is_exist(web_driver, locator, timeout=timeout)
    if element_exist:
        ele = web_driver.find_element(By.XPATH, locator)
        if clear_flag:
            ele.clear()
        ele.send_keys(text)
    else:
        raise TimeoutException(f"查找元素：{locator}超时.")


def hover_element(web_driver: WebDriver, locator: str, timeout: float = 10):
    element_exist = element_is_exist(web_driver, locator, timeout=timeout)
    if element_exist:
        ele = web_driver.find_element(By.XPATH, locator)
        ActionChains(web_driver).move_to_element(ele).perform()
        time.sleep(1)
    else:
        raise TimeoutException(f"查找元素：{locator}超时.")


def go_back(web_driver: WebDriver):
    """
    页面后退
    浏览器操作：后退
    """
    web_driver.back()


def reload_page(web_driver: WebDriver):
    """
    刷新页面
    刷新当前页面
    """
    web_driver.refresh()


def execute_javascript(web_driver: WebDriver, js_code):
    web_driver.execute_script(js_code)


def get_web_info(web_driver: WebDriver, attr_name:str):
    if attr_name == "标题":
        return web_driver.title
    elif attr_name == "html源码":
        return web_driver.page_source
    elif attr_name == "网址":
        return web_driver.current_url
    elif attr_name == "cookies":
        return web_driver.get_cookies()
    else:
        return "不支持的网页信息。"


def get_element_attr(web_driver: WebDriver, locator: str, attr_name: str, timeout: float = 10):
    element_exist = element_is_exist(web_driver, locator, timeout=timeout)
    if element_exist:
        ele = web_driver.find_element(By.XPATH, locator)
        if attr_name == "文本":
            return ele.text
        elif attr_name == "值":
            return ele.get_attribute("value")
        else:
            return ele.get_attribute(attr_name)
    else:
        raise TimeoutException(f"查找元素：{locator}超时.")


def handle_alert(web_driver: WebDriver):
    try:
        alert = web_driver.switch_to.alert
        alert_text = alert.text
        alert.accept()
        return alert_text
    except NoAlertPresentException as e:
        return None


def switch_iframe(web_driver: WebDriver, locator: str, timeout: float = 10):
    element_exist = element_is_exist(web_driver, locator, timeout=timeout)
    if element_exist:
        ele = web_driver.find_element(By.XPATH, locator)
        web_driver.switch_to.frame(ele)
    else:
        raise TimeoutException(f"查找iframe：{locator}超时.")


def exit_iframe(web_driver: WebDriver, timeout: float = 10):
    web_driver.switch_to.default_content()


def switch_window(web_driver: WebDriver, window: Union[str, int] = "", switch_type: str = "标题", timeout: float = 10):
    begin_time = time.time()
    used_time = time.time()
    while used_time - begin_time < timeout:
        try:
            window_handles = web_driver.window_handles
            if switch_type == "最新":
                handle = window_handles[-1]
                web_driver.switch_to.window(handle)
            elif switch_type == "位置":
                if window > len(window_handles):
                    raise IndexError("窗口序号大于总窗口数量。")
                handle = window_handles[window-1]
                web_driver.switch_to.window(handle)
            elif switch_type == "标题":
                for window_handle in window_handles:
                    web_driver.switch_to.window(window_handle)
                    window_title = web_driver.title
                    if window_title == window:
                        break
                else:
                    raise NoSuchWindowException(f"未找到该窗口：{window}")
            else:
                pass
            break
        except Exception as e:
            used_time = time.time()
    else:
        window_handles = web_driver.window_handles
        if switch_type == "最新":
            handle = window_handles[-1]
            web_driver.switch_to.window(handle)
        elif switch_type == "位置":
            if window > len(window_handles):
                raise IndexError("窗口序号大于总窗口数量。")
            handle = window_handles[window - 1]
            web_driver.switch_to.window(handle)
        elif switch_type == "标题":
            for window_handle in window_handles:
                web_driver.switch_to.window(window_handle)
                window_title = web_driver.title
                if window_title == window:
                    break
            else:
                raise NoSuchWindowException(f"未找到该窗口：{window}")
        else:
            pass


#
# def element(web_driver: WebDriver, locator: str, timeout: float = 10):
#     element_exist = element_is_exist(web_driver, locator, timeout=timeout)
#     if element_exist:
#         ele = web_driver.find_element(By.XPATH, locator)
#
#     else:
#         raise TimeoutException(f"查找元素：{locator}超时.")


if __name__ == '__main__':
    driver = open_browser()
    open_url(driver, "www.baidu.com")










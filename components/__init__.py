"""

By Ziqiu Li
Created at 2023/2/15 14:23
"""

__version__ = "0.0.0"

import re
import os
import requests
import winreg
import zipfile
from difflib import SequenceMatcher


def get_chrome_version():
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Google\Chrome\BLBeacon')
        version, _ = winreg.QueryValueEx(key, 'version')
        return version
    except Exception as e:
        # print(1)
        raise FileNotFoundError("未找到Chrome版本号，请检查是否安装Chrome浏览器.")


def get_chrome_driver_version():
    version_text = os.popen('chromedriver --version').read()
    version = ""
    if re.findall(r"ChromeDriver (\d+\.\d+\.\d+\.\d+) ", version_text):
        version = re.findall(r"ChromeDriver (\d+\.\d+\.\d+\.\d+) ", version_text)[0]
    return version


def get_chrome_driver_version_list():
    url = "https://registry.npmmirror.com/-/binary/chromedriver/"
    r = requests.get(url)
    version_list = []
    for row_data in r.json():
        if row_data["type"] == "dir":
            version_list.append(row_data["name"].replace("/", ""))
    return version_list


def get_similarity(str1, str2):
    return SequenceMatcher(None, str1, str2).ratio()


def unzip_driver_file():
    f = zipfile.ZipFile("chromedriver.zip", "r")
    for file in f.namelist():
        f.extract(file)


def download_chrome_driver(version):
    download_url = f"https://registry.npmmirror.com/-/binary/chromedriver/{version}/chromedriver_win32.zip"
    r = requests.get(download_url)
    with open("chromedriver.zip", "wb") as f:
        f.write(r.content)


def check_chrome_driver():
    chrome_version = get_chrome_version()
    chrome_driver_version = get_chrome_driver_version()
    if chrome_version != chrome_driver_version:
        driver_list = get_chrome_driver_version_list()
        if chrome_version in driver_list:
            version = chrome_version
        else:
            similarity_list = list(map(lambda x: get_similarity(chrome_version, x), driver_list))
            if max(similarity_list) > 5:
                version = driver_list[similarity_list.index(max(similarity_list))]
            else:
                raise FileNotFoundError("未找到匹配的ChromeDriver版本.")
        if version == chrome_driver_version:
            print(f"当前Chrome版本：{chrome_version}, 当前ChromeDriver版本: {chrome_driver_version}为匹配版本, 无需更新.")
        else:
            download_chrome_driver(version)
            unzip_driver_file()
            print(f"ChromeDriver版本已下载/更新, Chrome版本: {chrome_version}, ChromeDriver版本：{version}")
    else:
        print("Chrome与ChromeDriver版本相同, 无需下载/更新.")


if __name__ == "__main__":
    check_chrome_driver()
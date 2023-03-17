"""
Setup everything,
and then build with `python -m setup sdist bdist_wheel`,
and then upload with `twine upload --repository-url https://upload.pypi.org/legacy/ dist/*`

By Allen Tao
Created at 2023/3/17 16:05
"""
import setuptools

import qiuying

with open("README.md", "r") as f:
    long_description = f.read()
    setuptools.setup(
        name="qiuying",
        version=qiuying.__version__,
        author="Allen Tao, Ziqiu Li",
        author_email="allen@tkzt.cn, lcmail1001@163.com",
        description="A plain, simple and this-is-the-way library for RPA developing.",
        long_description=long_description,
        long_description_content_type="text/markdown",
        url="https://github.com/qiuying-rpa/qiuying",
        packages=setuptools.find_packages(exclude=['tests', 'README.md']),
        classifiers=[
            "Programming Language :: Python :: 3",
            "License :: OSI Approved :: MIT License",
            "Operating System :: Microsoft :: Windows",
        ],
        install_requires=['pynput>=1.7.6', 'pywin32>=305', 'selenium>=4.8.2', 'requests>=2.28.2']
    )

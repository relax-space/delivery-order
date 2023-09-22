from setuptools import setup, find_packages
import os


def package_files(directory):
    paths = []
    for path, directories, filenames in os.walk(directory):
        for filename in filenames:
            paths.append(os.path.join("..", path, filename))
    return paths


extra_files = package_files("base")
setup(
    name="main",
    version="1.0.0",
    author="小苗",
    author_email="your@email.com",
    description="送货单",
    packages=find_packages(),
    include_package_data=True,  # 包括数据文件
    package_data={
        "": extra_files,  # 指定包含的数据文件
    },
)

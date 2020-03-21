import setuptools


setuptools.setup(
    name="genquiz",
    version="1.0.4",
    author="Parinya Sanguansat",
    author_email="sanguansat@yahoo.com",
    description="Create individual questions for online examination",
    long_description="Create individual questions for online examination",
    long_description_content_type="text/markdown",
    url="https://github.com/Parinyasan/genquiz",
    packages=setuptools.find_packages(),
    classifiers=(
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: GNU General Public License v3 or later (GPLv3+)",
        "Operating System :: OS Independent",
    ),
    install_requires=['openpyxl', 'numpy', 'python-docx', 'lxml', 'sympy', 'pandas', 'xlrd'],
    include_package_data=True
)

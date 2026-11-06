from setuptools import setup, find_packages

setup(
    name="touyan2",
    version="1.1.0",
    packages=find_packages(),
    install_requires=["python-docx>=0.8.11"],
    entry_points={
        "console_scripts": [
            "touyan2 = touyan2.cli:main",
        ],
    },
)

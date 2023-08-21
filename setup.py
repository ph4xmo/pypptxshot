import re

from setuptools import setup
from setuptools.extern import packaging

# Get version to show the same as twine.
with open("pypptxshot/__init__.py", "rt", encoding="utf-8") as f:
    version_re = re.search(r"__version__ = \"(.*?)\"", f.read())
    if version_re:
        version = version_re.group(1)
    else:
        raise ValueError("Could not determine package version")
    version = str(packaging.version.Version(version))

with open("README.md", "r", encoding="utf-8") as readme_file:
    long_description = readme_file.read()

setup(
    name="PyPPTXShot",
    version=version,
    description="A Python Screenshotting PPTX Generator",
    long_description=long_description,
    long_description_content_type="text/markdown",
    author="ph4xmo",
    author_email="ph4xmo@gmail.com",
    license="MIT",
    url="https://github.com/ph4xmo/pypptxshot/",
    classifiers=[
        "Intended Audience :: Developers",
        "Intended Audience :: End Users/Desktop",
        "Intended Audience :: Education",
        "License :: OSI Approved :: MIT License",
        "Topic :: Software Development",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
    ],
    keywords="clicker autoclicker pptx ppt generator screenshot",
    packages=["pypptxshot"],
    install_requires=[
        "PyAutoGUI==0.9.54",
        "python_pptx==0.6.21",
    ],
    python_requires="~=3.8.10",
)

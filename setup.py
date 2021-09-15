"""Python setup.py for allin_hotels package"""
import io
import os
from setuptools import find_packages, setup


def read(*paths, **kwargs):
    """Read the contents of a text file safely.
    >>> read("allin_hotels", "VERSION")
    '0.1.0'
    >>> read("README.md")
    ...
    """

    content = ""
    with io.open(
        os.path.join(os.path.dirname(__file__), *paths),
        encoding=kwargs.get("encoding", "utf8"),
    ) as open_file:
        content = open_file.read().strip()
    return content


def read_requirements(path):
    return [
        line.strip()
        for line in read(path).split("\n")
        if not line.startswith(('"', "#", "-", "git+"))
    ]


setup(
    name="allin_hotels",
    version=read("allin_hotels", "VERSION"),
    description="Awesome allin_hotels created by zhou322",
    url="https://github.com/zhou322/allin-hotels/",
    long_description=read("README.md"),
    long_description_content_type="text/markdown",
    author="zhou322",
    packages=find_packages(exclude=["tests", ".github"]),
    install_requires=read_requirements("requirements.txt"),
    entry_points={
        "console_scripts": ["allin_hotels = allin_hotels.__main__:main"]
    },
    extras_require={"test": read_requirements("requirements-test.txt")},
)

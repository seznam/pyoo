"""
PyOO - Pythonic interface to Apache OpenOffice API (UNO)

Copyright (c) 2014 Seznam.cz, a.s.

"""


from distutils.core import setup

setup(
    name='pyoo',
    version='0.4',
    description='Pythonic interface to OpenOffice.org API (UNO)',
    long_description = open('README.rst').read(),
    author='Miloslav Pojman',
    author_email='miloslav.pojman@firma.seznam.cz',
    py_modules=['pyoo'],
)

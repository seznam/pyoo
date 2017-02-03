"""
PyOO - Pythonic interface to Apache OpenOffice API (UNO)

Copyright (c) 2016 Seznam.cz, a.s.

"""


from distutils.core import setup

setup(
    name='pyoo',
    version='1.3.dev',
    description='Pythonic interface to Apache OpenOffice API (UNO)',
    long_description = open('README.rst').read(),
    author='Miloslav Pojman',
    author_email='miloslav.pojman@firma.seznam.cz',
    url='https://github.com/seznam/pyoo',
    py_modules=['pyoo'],
    classifiers=[
        'Intended Audience :: Developers',
        'License :: OSI Approved :: MIT License',
        'Programming Language :: Python',
        'Topic :: Office/Business :: Office Suites',
        'Topic :: Office/Business :: Financial :: Spreadsheet',
    ],
)

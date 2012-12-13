
from distutils.core import setup

setup(
    name='pyoo',
    version='0.1',
    description='Pythonic interface to OpenOffice.org API known as UNO.',
    long_description = open('README.rst').read(),
    author='Miloslav Pojman',
    author_email='miloslav.pojman@firma.seznam.cz',
    url='http://cml.kancelar.seznam.cz/generic/browser/szn-python-pyoo',
    py_modules=['pyoo'],
)
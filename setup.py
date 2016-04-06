"""
This is a setup.py script generated by py2applet

Usage:
    python setup.py py2app
"""

from setuptools import setup

APP = ['app.py']
DATA_FILES = []
OPTIONS = {'argv_emulation': True,
           'includes': ['sip', 'PyQt4', 'PyQt4.QtCore', 'PyQt4.QtGui'],
           'iconfile': 't2xl.icns'
           }

setup(
    name='CatDV to XLSX',
    version='2.0.1',
    description='Convert CatDV .txt output to .xlsx',
    date='3-Dec-2015',
    url='https://github.com/edsoncudjoe/CatDVText2XlsxGui',

    author='Edson Cudjoe',
    author_email='bashpythonstuff@hotmail.co.uk',
    license='MIT',

    classifiers=[
        'Development Status :: 4 - Beta',
        'Intended Audience :: Developers',
        'Topic :: Software Development :: Media',
        'License :: OSI Approved :: MIT License',
        'Programming Language :: Python :: 2',
        'Programming Language :: Python :: 2.6',
        'Programming Language :: Python :: 2.7',
    ],

    keywords='catdv text xlsx',

    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)

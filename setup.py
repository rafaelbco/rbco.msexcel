from setuptools import setup, find_packages
import os

version = '0.0.2'

setup(name='rbco.msexcel',
      version=version,
      description='Provide functions to read, parse and convert MS Excel '
        'spreadsheets into various data structures.',
      long_description=open(os.path.join('rbco', 'msexcel', 'README.txt')).read(),
      # Get more strings from http://www.python.org/pypi?%3Aaction=list_classifiers
      classifiers=[
        "Programming Language :: Python",
        "Topic :: Software Development :: Libraries :: Python Modules",
        ],
      keywords='xls excel',
      author='Rafael Oliveira',
      author_email='rafaelbco@gmail.com',
      url='',
      license='GPL',
      packages=find_packages(exclude=['ez_setup']),
      namespace_packages=['rbco'],
      include_package_data=True,
      zip_safe=False,
      install_requires=[
          'setuptools',
          'pyExcelerator>=0.6.4',
      ],
      entry_points="""
      # -*- Entry points: -*-
      """,
)

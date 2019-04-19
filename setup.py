#!/usr/bin/env python
# encoding: utf-8

from setuptools import setup, find_packages

setup(
    name='advanced_tools',
    version='0.1',
    description='Utility functions,scripts for daily operations',
    url='http://github.com/kvdogan/advanced_tools',
    author='kvdogan',
    packages=find_packages(),
    include_package_data=True,
    license='MIT License',
    entry_points={
        'console_scripts': ['build-hierarchy=advanced_tools.build_hierarchy_tree:main'],
      }
)


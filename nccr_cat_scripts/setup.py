#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Nov 24 17:42:16 2025

@author: Niccolò Ricardi
"""

from setuptools import setup, find_packages
VERSION = '0.1.0'

setup(
    name='nccr_cat_scripts',
    version=VERSION,
    
    # find_packages() automatically detects the 'nccr_cat_scripts' directory 
    # as a package because it contains an __init__.py. This is scalable.
    packages=find_packages(),
    
    url='https://github.com/useful_scripts', # Update this with the specific URL
    license='GPL-3.0',
    author='Niccolò Ricardi',
    author_email='Niccolo.Ricardi@epfl.ch',
    
    # Hardcoded short description
    description='A collection of useful utilities for the general NCCR Catalysis community.',
    
    # Hardcoded long description (as requested, avoiding file read)
    long_description='A collection of useful utilities for the general NCCR Catalysis community. This package provides both CLI commands and importable functions.',
    long_description_content_type='text/plain', 
    
    # No external dependencies are currently required for zip_utils.py
    install_requires=[],

    # This section defines the command-line executables
    entry_points={
        'console_scripts': [
            # The CLI command will be 'nccr-zip'. 
            # It executes the 'cli' function inside 'nccr_cat_scripts.zip_utils'.
            'zip-utils = nccr_cat_scripts.zip_utils:cli',
        ],
    },
    
    # Metadata for PyPI/packaging
    classifiers=[
        'Development Status :: 3 - Alpha',
        'Intended Audience :: Developers',
        'Topic :: System :: Archiving :: Compression',
        'License :: OSI Approved :: GNU General Public License v3 or later (GPLv3+)', # Updated to GPL-3.0
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.6',
        'Programming Language :: Python :: 3.7',
        'Programming Language :: Python :: 3.8',
        'Programming Language :: Python :: 3.9',
        'Programming Language :: Python :: 3.10',
        'Programming Language :: Python :: 3.11',
    ],
    python_requires='>=3.6',
)
#!/usr/bin/python
# -*- coding: utf-8 -*-
# Copyright 2019 Shieber 
# All Rights Reserved.
#
#    Licensed under the Apache License, Version 2.0 (the "License"); you may
#    not use this file except in compliance with the License. You may obtain
#    a copy of the License at
#
#         http://www.apache.org/licenses/LICENSE-2.0
#
#    Unless required by applicable law or agreed to in writing, software
#    distributed under the License is distributed on an "AS IS" BASIS, WITHOUT
#    WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the
#    License for the specific language governing permissions and limitations
#    under the License.

import os
import re
import codecs

from setuptools import setup

here = os.path.abspath(os.path.dirname(__file__))


# Read the version number from a source file.
def find_version(*file_paths):
    # Open in Latin-1 so that to avoid encoding errors.
    # Use codecs.open for Python 2 compatibility
    try:
        f = codecs.open(os.path.join(here, *file_paths), 'r', 'latin1')
        version_file = f.read()
        f.close()
    except:
        raise RuntimeError("Unable to find version string.")

    # The version line must have the form  __version__ = 'number'
    version_match = re.search(r"^__version__ = ['\"]([^'\"]*)['\"]",
                              version_file, re.M)
    if version_match:
        return version_match.group(1)
    raise RuntimeError("Unable to find version string.")


# Get the long description from the relevant file
try:
    f = codecs.open('README.rst', encoding='utf-8')
    long_description = f.read()
    f.close()
except:
    long_description = ''


setup(
    name='Billstat',
    version=find_version('Billstat.py'),
a cron-job to statistics your expendirture.
    description=('a cron-job to statistics your expendirture. '
                 'Billstat'),
    long_description=long_description,
    keywords='Billstat',
    author='Shieber',
    author_email='QMH_XB_FLTMY@yahoo.com',
    url='https://github.com/QMH-TMY/Billstat',
    license='Apache License, Version 2.0',
    py_modules=['Billstat'],
    entry_points={
        'console_scripts': ['=Billstat:main']
    },
    classifiers=[
        'Development Status :: 5 - Production/Stable',
        'Programming Language :: Python',
        'Environment :: Console',
        'License :: OSI Approved :: Apache Software License',
        'Operating System :: OS Independent',
        'Programming Language :: Python :: 2.7',
    ]
)

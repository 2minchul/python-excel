#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from setuptools import setup
import sys



setup(
    name="excel",
    version="0.4",
    install_requires=[
        "openpyxl", "xlrd"
    ],
    packages=['excel'],
)

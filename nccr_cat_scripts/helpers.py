#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Jan  7 16:29:52 2026

@author: nr
"""

from collections.abc import Collection
import os


def islistlike(obj):
    return isinstance(obj, Collection) and not isinstance(obj, (str, bytes))

def isfile(path):
    if os.path.exists(path):
        return os.path.isfile(path)
    else:
        return bool(os.path.splitext(path)[1])

def isdir(path):
    if os.path.exists(path):
        return os.path.isdir(path)
    else:
        return not bool(os.path.splitext(path)[1])

def check_and_clean_folderpath(path):
    assert isdir(path), f"It looks like you provided a filepath ({path}), while the code was expecting a folder path."
    if not path.endswith(os.path.sep):
        path = f"{path}{os.path.sep}"
    return path

def harmonize_ext(ext):
    if ext.startswith("."):
        return ext[1:]
    return ext

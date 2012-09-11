#! /usr/bin/env python
# -*- coding: utf-8 -*-

import os
import glob
import time
import subprocess

from datetime import datetime, date

def type_and_value(val):
    _type = type(val)
    if _type in [str, int, long, float]:
        return None, val
    elif _type is datetime:
        return "datetime", time.mktime(val.timetuple())
    elif _type is date:
        return "date", time.mktime(val.timetuple())
    elif _type is bool:
        return None, str(val)
    else:
        return None, val

def autoload():
    here = os.path.abspath(os.path.dirname(__file__))

    suggestions = glob.glob(os.path.join(here, "scaxls-*.jar"))
    if len(suggestions) == 0:
        raise RuntimeError("library not found.")

    return sorted(suggestions, reverse=True)[0]

def json2xls(json, libloader=autoload):
    proc = subprocess.Popen("java -jar {0}".format(libloader()),
                            shell=True,
                            stdin=subprocess.PIPE,
                            stdout=subprocess.PIPE,
                            stderr=subprocess.PIPE)

    out, err = proc.communicate(json)
    return out

def xls2json(filename, libloader=autoload):
    proc = subprocess.Popen("java -jar {0} -template {1} -mode read".format(libloader(), filename),
                            shell=True,
                            stdout=subprocess.PIPE,
                            stderr=subprocess.PIPE)
    out, err = proc.communicate()
    return out

if __name__ == '__main__':
    import sys
    #print xls2json(sys.argv[1])

    print json2xls(sys.stdin.read())

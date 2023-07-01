# 基于python的libreoffice工具.

## 设定 console for libreoffice.

```
# -*- coding: utf-8 -*-

from __future__ import unicode_literals

import uno, os, subprocess

def interpreter_console():
    ctx = XSCRIPTCONTEXT.getComponentContext() 
    smgr = ctx.getServiceManager()
    ps = smgr.createInstanceWithContext("com.sun.star.util.PathSettings", ctx)
    install_path = uno.fileUrlToSystemPath(ps.Module)
    pgm = install_path + os.sep + "python"  # Python shell/console path
    subprocess.Popen(pgm)  # Start Python interactive Shell
```

参考 > https://help.libreoffice.org/latest/en-US/text/sbasic/python/python_shell.html?DbPAR=BASIC#N0117

## 测试marco第一个python脚本.

```
#!/bin/python3

import uno

def my_first_macro_calc():
    doc = XSCRIPTCONTEXT.getDocument()
    try :
        cell = doc.Sheets['test']['A1']  # com.sun.star.sheet.XSpreadsheetDocument
        cell.setString('Hello World in Python in Calc')
    except :
        cell = doc.Sheets[0]['A1']
        cell.setString()
    return
```
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

#! /usr/bin/python3

import sys
import uno

local = uno.getComponentContext()
resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
context = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
desktop = context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", context)

url = "private:factory/swriter"
document = desktop.loadComponentFromURL(url, "_blank", 0, ())
cursor = document.Text.createTextCursor()

text = "Fuck this sheet.\n\n"
document.Text.insertString(cursor, text, 0)

document.storeAsURL("file:///home/canis/001.odt",())
document.dispose()

sys.exit()

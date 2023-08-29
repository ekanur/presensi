import win32print

printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL, None, 1)
print (printers)
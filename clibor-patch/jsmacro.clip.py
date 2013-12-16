#
# -*- coding: utf-8 -*-
#
# Clibor JavaScript Macro Clipboard Helper
#
import sys
import codecs
import win32api
import win32com.client
import win32clipboard as CB

params = sys.argv
tmpclip = params[1].replace(r'[\x5C]', '/')
mode = params[2]

win32api.Sleep(10)
shell = win32com.client.Dispatch("WScript.Shell")

if mode == "get":
  CB.OpenClipboard()
  data = CB.GetClipboardData(CB.CF_UNICODETEXT)
  CB.CloseClipboard()

  f = codecs.open(tmpclip, "w", "UTF-8")
  f.write(data)
  f.close()

else:
  f = codecs.open(tmpclip, "r", "UTF-8")
  text = f.read()
  f.close()

  CB.OpenClipboard()
  CB.EmptyClipboard()
  CB.SetClipboardText(text, CB.CF_UNICODETEXT)
  CB.CloseClipboard()


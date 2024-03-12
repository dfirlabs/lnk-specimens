#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Script to generate Windows Shell link test files.

Requires Windows and pywin32.
"""

import os

import pythoncom

from win32com import storagecon
from win32com.propsys import propsys
from win32com.shell import shell


if __name__ == '__main__':
  specimens_path = os.path.join(os.getcwd(), 'specimens')

  os.makedirs(specimens_path)

  # Shell link with path.
  shortcut = pythoncom.CoCreateInstance(
      shell.CLSID_ShellLink, None, pythoncom.CLSCTX_INPROC_SERVER,
      shell.IID_IShellLink)

  shortcut.SetPath('C:\\test')

  interface = shortcut.QueryInterface(pythoncom.IID_IPersistFile)

  path = os.path.join(specimens_path, 'path.lnk')
  interface.Save(path, 0)

  # Shell link with properties.
  shortcut = pythoncom.CoCreateInstance(
      shell.CLSID_ShellLink, None, pythoncom.CLSCTX_INPROC_SERVER,
      shell.IID_IShellLink)

  shortcut.SetPath('C:\\test')

  property_store = shortcut.QueryInterface(propsys.IID_IPropertyStore)

  property_key = propsys.PSGetPropertyKeyFromName('System.Document.LastAuthor')
  property_value = propsys.PROPVARIANTType(None, pythoncom.VT_NULL)
  property_store.SetValue(property_key, property_value)

  property_key = propsys.PSGetPropertyKeyFromName('System.Image.ResolutionUnit')
  property_value = propsys.PROPVARIANTType(0x1234, pythoncom.VT_I2)
  property_store.SetValue(property_key, property_value)

  property_key = propsys.PSGetPropertyKeyFromName('System.Message.ToDoFlags')
  property_value = propsys.PROPVARIANTType(0x12345678, pythoncom.VT_I4)
  property_store.SetValue(property_key, property_value)

  property_key = propsys.PSGetPropertyKeyFromName('System.GPS.DOP')
  property_value = propsys.PROPVARIANTType(1.2345, pythoncom.VT_R4)
  property_store.SetValue(property_key, property_value)

  property_key = propsys.PSGetPropertyKeyFromName('System.Photo.MaxAperture')
  property_value = propsys.PROPVARIANTType(1.2345, pythoncom.VT_R8)
  property_store.SetValue(property_key, property_value)

  # VT_CY
  # VT_DATE

  property_key = propsys.PSGetPropertyKeyFromName('System.Title')
  property_value = propsys.PROPVARIANTType('My Title', pythoncom.VT_BSTR)
  property_store.SetValue(property_key, property_value)

  # VT_ERROR

  property_key = propsys.PSGetPropertyKeyFromName(
      'System.Search.IsClosedDirectory')
  property_value = propsys.PROPVARIANTType(True, pythoncom.VT_BOOL)
  property_store.SetValue(property_key, property_value)

  # VT_VARIANT

  # ValueError: argument is not a COM object (got type=str)
  # property_key = propsys.PSGetPropertyKeyFromName('System.Devices.NotificationStore')
  # property_value = propsys.PROPVARIANTType('My unknown', pythoncom.VT_UNKNOWN)
  # property_store.SetValue(property_key, property_value)

  # VT_DECIMAL

  # VT_I1

  property_key = propsys.PSGetPropertyKeyFromName('System.Photo.Flash')
  property_value = propsys.PROPVARIANTType(0x12, pythoncom.VT_UI1)
  property_store.SetValue(property_key, property_value)

  property_key = propsys.PSGetPropertyKeyFromName('System.Image.ColorSpace')
  property_value = propsys.PROPVARIANTType(0x1234, pythoncom.VT_UI2)
  property_store.SetValue(property_key, property_value)

  property_key = propsys.PSGetPropertyKeyFromName(
      'System.Photo.DigitalZoomNumerator')
  property_value = propsys.PROPVARIANTType(0x12345678, pythoncom.VT_UI4)
  property_store.SetValue(property_key, property_value)

  # VT_I8

  property_key = propsys.PSGetPropertyKeyFromName('System.FileCount')
  property_value = propsys.PROPVARIANTType(0x123456789abcdef0, pythoncom.VT_UI8)
  property_store.SetValue(property_key, property_value)

  # VT_INT
  # VT_UINT
  # VT_VOID
  # VT_HRESULT
  # VT_PTR
  # VT_SAFEARRAY
  # VT_CARRAY
  # VT_USERDEFINED

  # TypeError: Unsupported property type 0x1e
  # property_key = propsys.PSGetPropertyKeyFromName('System.Comment')
  # property_value = propsys.PROPVARIANTType('My Comment', pythoncom.VT_LPSTR)
  # property_store.SetValue(property_key, property_value)

  property_key = propsys.PSGetPropertyKeyFromName('System.ItemNameDisplay')
  property_value = propsys.PROPVARIANTType('My Item', pythoncom.VT_LPWSTR)
  property_store.SetValue(property_key, property_value)

  # VT_RECORD
  # VT_INT_PTR
  # VT_UINT_PTR

  # TypeError: must be a pywintypes time object (got int)
  # property_key = propsys.PSGetPropertyKeyFromName('System.Search.GatherTime')
  # property_value = propsys.PROPVARIANTType(0x123456789abcdef0, pythoncom.VT_FILETIME)
  # property_store.SetValue(property_key, property_value)

  property_key = propsys.PSGetPropertyKeyFromName(
      'System.Music.SynchronizedLyrics')
  property_value = propsys.PROPVARIANTType(b'My BLOB', pythoncom.VT_BLOB)
  property_store.SetValue(property_key, property_value)

  # stream = pythoncom.CreateStreamOnHGlobal()
  # stream.Write(b'My Stream')

  flags = storagecon.STGM_READWRITE | storagecon.STGM_SHARE_EXCLUSIVE

  storage_path = os.path.join(specimens_path, 'com_storage')
  storage = pythoncom.StgCreateStorageEx(
      storage_path, flags, storagecon.STGFMT_STORAGE, 0, pythoncom.IID_IStorage)

  stream = storage.CreateStream('com_stream', flags)

  # a VT_STREAM property key does appears to be stored empty
  # possibly related to "simple property set"?
  property_key = propsys.PSGetPropertyKeyFromName('System.ThumbnailStream')
  property_value = propsys.PROPVARIANTType(stream, pythoncom.VT_STREAM)
  property_store.SetValue(property_key, property_value)

  # a VT_STORAGE property key does not appear to be stored
  # possibly related to "simple property set"?
  property_key = propsys.PSGetPropertyKeyFromName('System.Photo.MakerNote')
  property_value = propsys.PROPVARIANTType(storage, pythoncom.VT_STORAGE)
  property_store.SetValue(property_key, property_value)

  # VT_STREAMED_OBJECT
  # VT_STORED_OBJECT
  # VT_BLOB_OBJECT

  # TypeError: Unsupported property type 0x47
  # property_key = propsys.PSGetPropertyKeyFromName('System.Thumbnail')
  # property_value = propsys.PROPVARIANTType(b'My Clipboard', pythoncom.VT_CF)
  # property_store.SetValue(property_key, property_value)

  # Note the required string format is "{%GUID%}"
  property_key = propsys.PSGetPropertyKeyFromName('System.NamespaceCLSID')
  property_value = propsys.PROPVARIANTType(
      '{525BF964-FF68-4D52-9E1D-AED5051FA555}', pythoncom.VT_CLSID)
  property_store.SetValue(property_key, property_value)

  # VT_VERSIONED_STREAM

  interface = shortcut.QueryInterface(pythoncom.IID_IPersistFile)

  path = os.path.join(specimens_path, 'properties.lnk')
  interface.Save(path, 0)

#!/usr/bin/env python
# -- coding: utf-8 --
# @Author : wxy
# @File : MSIFile.py
# Ported to Python 3, translated, Hm
# Todo: Change the very last append to the real length

import ctypes
msidll = ctypes.windll.msi

def storeHandle(func):
    """Decorator function for manipulating storage handles and pointers"""
    def store(self, *args, **kwargs):
        Handle, pHandle = func(self, *args, **kwargs)
        self.__handle_list__.append(Handle)
        self.__phandle_list__.append(pHandle)
        return Handle, pHandle
    return store



class MsiFile:
    """
    Refer to msiquery.h, msi.h use Ctypes to rewrite WinAPI for manipulating MSI files
    For reasons of demand, it only involves reading information. For information about writing msi files, please refer to pypi-msilib
    """

    def __init__(self, file_path, sql):

        #The following list holds all open handles and corresponding pointers
        self.__handle_list__ = []
        self.__phandle_list__ = []

        #According to Microsoft documentation, open the database and execute SQL statement query information step by step
        try:
            self.hDataBase, self.phDataBase = self.OpenDataBase(file_path)
            self.hView, self.phView = self.DatabaseOpenViewW(sql)
            self.hRecord, self.phRecord = self.Execute()
        except:
            raise
        # The result must be fetched by Fetch before it can get the result. For multiple rows of results, Fetch fetches one row of results at a time.

    def close(self):
        """The object calls close with end to close the handle"""
        for handle in self.__handle_list__:
            msidll.MsiCloseHandle(handle)

    def __del__(self):
        self.close()
    @storeHandle
    def OpenDataBase(self, file_path):
        hDataBase = ctypes.c_ulong(0)
        phDataBase = ctypes.pointer(hDataBase)
        file_path = ctypes.c_char_p(file_path.encode("ascii"))
        MSI_READ_ONLY = "1"
        szPersist = ctypes.c_char_p(MSI_READ_ONLY.encode("ascii")) # MSI_READ_ONLY
        if msidll.MsiOpenDatabaseA(file_path, szPersist, phDataBase) != 0:
            msidll.MsiCloseHandle(hDataBase)
            print("Failed to open database handle, check file path and file is in MSI format")
            raise
        return hDataBase, phDataBase

    @storeHandle
    def DatabaseOpenViewW(self, sql):
        psql = ctypes.c_char_p(sql.encode("ascii"))
        hView = ctypes.c_ulong(0)
        phView = ctypes.pointer(hView)
        if msidll.MsiDatabaseOpenViewA(self.hDataBase, psql, phView) != 0:
            msidll.MsiCloseHandle(hView)
            print("Failed to open the View object, check whether the SQL statement is correct")
            raise
        return hView, phView

    @storeHandle
    def Execute(self):
        hRecord = ctypes.c_ulong(0)
        msidll.MsiViewExecute(self.hView, hRecord)
        phRecord = ctypes.pointer(hRecord)
        return hRecord, phRecord

    def Fetch(self):
        """Fetch must be executed after executing Execute to obtain data"""
        msidll.MsiViewFetch(self.hView, self.phRecord)

    def RecordGetFieldCount(self):
        """Returns the number of entries, that is, the number of columns in the data table"""
        counts  = msidll.MsiRecordGetFieldCount(self.hRecord)
        return counts

    def RecordGetString(self,prev_string_length=8,field=1):
        """Read string"""
        field = ctypes.c_uint(field)
        buf = (ctypes.c_char * prev_string_length)()
        pbuf = ctypes.pointer(buf)
        length = ctypes.c_ulong(prev_string_length)
        plength  = ctypes.pointer(length)
        msidll.MsiRecordGetStringA(self.hRecord, field, pbuf, plength)
        return buf

    def ReadStream(self, prev_stream_size=2048, field=1):
        """It is useful to read binary data, the default size is 2048 each time."""
        field = ctypes.c_uint(field)
        buf = (ctypes.c_char * prev_stream_size)()
        pbuf = ctypes.pointer(buf)
        buf_size = ctypes.sizeof(buf)
        flag_size = prev_stream_size
        buf_size = ctypes.c_long(buf_size)
        pbuf_size = ctypes.pointer(buf_size)
        status = 0
        res = []
        while status == 0:
            status = msidll.MsiRecordReadStream(
            self.hRecord, field, pbuf, pbuf_size)
            res.append(buf.raw)
            # Todo: Change the very last append to the real length
            # print(buf.raw)
            if buf_size.value != flag_size:
                break
        data = b''.join(res)
        
        return data

#The following examples are used to read the data in the MSI database
def getIconData(file_path):
    """
    example: sql = "select Data  From Icon"
    Get ICON data, possibly in PE format and ICON format. Subsequent judgment to use
     Reference: https://docs.microsoft.com/en-us/windows/win32/msi/icon-table
    """
    sql = "select Data From Icon"
    myMsiFile = MsiFile(file_path, sql)
    myMsiFile.Fetch()    #At present, it is known that the ICON table has only one row of data, so it is only Fetch once, and there is an error, and then change it
    if myMsiFile.RecordGetFieldCount() >= 1:
        buf = myMsiFile.ReadStream(prev_stream_size=2048, field=1)
        return buf
    else:
        return None

if __name__ == "__main__":
    file_path = 'C:\\tmp\\Test_with_Icon.msi'
    # example: sql = "select Data From Icon" # Get ICON data
    data = getIconData(file_path)
    binary_file = open("C:\\tmp\\Icon.exe", "wb")
    # Write bytes to file
    binary_file.write(data)
    # Close file
    binary_file.close()
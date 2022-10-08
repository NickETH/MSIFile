# MSIFile
Used to extract information from an MSI file. We use Ctypes to access the API for reading MSI file information in WINAPI.


# describe
Recently, needed to extract  an MSI ICON in Python 2.7. The Python MSIlib can only read String information, not Stream information. In fact, there is only one function missing, but I found that Ctypes can write WinAPI as long as it does not involve complex structures. Simple. So it was rewritten.

# Update to Python 3
Run the file trough 2to3.py and adjust some data type issues.
Todo: The length of the output is to long, because the last chunk of data is always the buffer size not actual length.




# MSIFile
Used to extract information from an MSI file。Use Ctypes to rewrite part of the API for reading MSI file information in WINAPI, which is convenient for calling.


# describe
Recently, Ptthon2.7 needs to extract MSI ICON. Using MSIlib can only read String information, not Stream information. In fact, there is only one function missing, but I found that Ctypes can write WinAPI as long as it does not involve complex structures. Simple. So it was rewritten.



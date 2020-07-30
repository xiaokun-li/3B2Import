from XML2SQL3B2 import *
import _mssql
from _mssql import decimal


"""
aa = parse3B2xml(r'c:\ftp\3b2.xml')
aa.extend(parse3B2xml(r'c:\ftp\3B2_IINTELSHCQ_00345.xml'))
for a in aa:
    print(a)

imp2db(aa)

"""


if __name__ == "__main__":

    #  1: 
    xmlfilelist = getxmlfiles("test")
    
    if not xmlfilelist:
        print("No file need to handle.")
        exit(status=0, message=None)
    else:
        for tmp in xmlfilelist:
            data = parse3B2xml(tmp)
            imp2db(data)

    bakxmlfromlist(xmlfilelist)

    print("finished.") 

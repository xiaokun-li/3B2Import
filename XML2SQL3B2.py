# Handle 3B2 XML and import the data into database

import xml.dom.minidom
import pymssql
import datetime
import os
import _mssql
from _mssql import decimal


#  .xml file to list
def parse3B2xml(_xml3b2):
    
    domtree = xml.dom.minidom.parse(_xml3b2)
    collection = domtree.documentElement
    result=[] 

    so = ''
    soline = ''
    dn = ''
    shiptocode = ''
    hppo = ''
    cdd = ''
    etd = ''
    sku = ''
    plantcode = ''
    qty = ''
    adddate=''
    reference1 = ''
    reference2 = ''

    so = collection.getElementsByTagName(r'ShippingContainerItem')[0].getElementsByTagName(r'ProprietaryDocumentIdentifier')[0].childNodes[0].data  
    soline = collection.getElementsByTagName(r'ShippingContainerItem')[0].getElementsByTagName(r'LineNumber')[0].childNodes[0].data
    dn = collection.getElementsByTagName(r'thisDocumentIdentifier')[0].getElementsByTagName(r'ProprietaryDocumentIdentifier')[0].childNodes[0].data
    shiptocode = collection.getElementsByTagName(r'BuyingPartner')[0].getElementsByTagName(r'GlobalLocationIdentifier')[0].childNodes[0].data
    hppo = collection.getElementsByTagName(r'ShippingContainerItem')[0].getElementsByTagName(r'ProprietaryDocumentIdentifier')[2].childNodes[0].data
    cdd = collection.getElementsByTagName(r'DateStamp')[2].childNodes[0].data
    etd = collection.getElementsByTagName(r'DateStamp')[1].childNodes[0].data
    plantcode = collection.getElementsByTagName(r'toRole')[0].getElementsByTagName(r'GlobalBusinessIdentifier')[0].childNodes[0].data
    reference2 = collection.getElementsByTagName(r'thisDocumentIdentifier')[0].getElementsByTagName(r'ProprietaryDocumentIdentifier')[0].childNodes[0].data
    shipmentitems  = collection.getElementsByTagName(r'ShippingContainerItem')

    for item in shipmentitems:
        
        sku = item.getElementsByTagName(r'ProprietaryProductIdentifier')[0].childNodes[0].data
        qty = item.getElementsByTagName(r'ProductQuantity')[6].childNodes[0].data
        reference1 = item.getElementsByTagName(r'ProprietaryLotIdentifier')[0].childNodes[0].data
     
        # so = collection.getElementsByTagName(r'ShippingContainerItem')[0].getElementsByTagName(r'ProprietaryDocumentIdentifier')[0].childNodes[0].data  
        # soline = collection.getElementsByTagName(r'ShippingContainerItem')[0].getElementsByTagName(r'LineNumber')[0].childNodes[0].data
        # dn = collection.getElementsByTagName(r'thisDocumentIdentifier')[0].getElementsByTagName(r'ProprietaryDocumentIdentifier')[0].childNodes[0].data
        # shiptocode = collection.getElementsByTagName(r'BuyingPartner')[0].getElementsByTagName(r'GlobalLocationIdentifier')[0].childNodes[0].data
        # hppo = collection.getElementsByTagName(r'ShippingContainerItem')[0].getElementsByTagName(r'ProprietaryDocumentIdentifier')[2].childNodes[0].data
        # cdd = collection.getElementsByTagName(r'DateStamp')[2].childNodes[0].data
        # etd = collection.getElementsByTagName(r'DateStamp')[1].childNodes[0].data
        # plantcode = collection.getElementsByTagName(r'toRole')[0].getElementsByTagName(r'GlobalBusinessIdentifier')[0].childNodes[0].data
        # reference2 = collection.getElementsByTagName(r'thisDocumentIdentifier')[0].getElementsByTagName(r'ProprietaryDocumentIdentifier')[0].childNodes[0].data

        recordline=[so,soline,dn,shiptocode,hppo,cdd,etd,sku,plantcode,qty,adddate,reference1,reference2]
        result.append(recordline)

    return result



def imp2db(_datalist):

    #  sqlserver database info:

    sqlhost = '10.213.27.231:1435'
    sqluser = 'wmsadmin'
    sqlpass = 'schenker123!@#'
    sqldatabase = 'wgq_hp'
    sqlstr="""insert into fd_intel_3b2(so,soline,dn,shiptocode,hppo,cdd,etd,sku,plantcode,qty,adddate,reference1,reference2) """

    conn = pymssql.connect(sqlhost, sqluser, sqlpass, sqldatabase)
    cursor = conn.cursor()

    so = ''
    soline = ''
    dn = ''
    shiptocode = ''
    hppo = ''
    cdd = ''
    etd = ''
    sku = ''
    plantcode = ''
    qty = ''
    adddate = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    reference1 = ''
    reference2 = ''

    for items in _datalist:
        
        sqlstr1=''

        so = items[0]
        soline = items[1]
        dn = items[2]
        shiptocode = items[3]
        hppo = items[4]
        cdd = items[5]
        etd = items[6]
        sku = items[7]
        plantcode = items[8]
        qty = items[9]
        # adddate = items[10]
        reference1 = items[11]
        reference2 = items[12]
    
        sqlstr1 = sqlstr + 'values(' +"'"+so+"'" +','+"'"+soline+"'"+','+"'"+dn+"'"+','+"'"+shiptocode+"'"\
            +','+"'"+hppo+"'"+','+"'"+cdd+"'"+','+"'"+etd+"'"+','+"'"+sku+"'"+','+"'"+plantcode+"'"+','+qty\
            +','+"'"+adddate+"'"+','+"'"+reference1+"'"+','+"'"+reference2+"'"+')'   # get datetime from python
        
        # sqlstr1 = sqlstr + 'values(' +"'"+so+"'" +','+"'"+soline+"'"+','+"'"+dn+"'"+','+"'"+shiptocode+"'"\
        #     +','+"'"+hppo+"'"+','+"'"+cdd+"'"+','+"'"+etd+"'"+','+"'"+sku+"'"+','+"'"+plantcode+"'"+','+qty\
        #     +','+'getdate()'+','+"'"+reference1+"'"+','+"'"+reference2+"'"+')'     #get datetime from sqlserver

        # print(sqlstr1)
        cursor.execute(sqlstr1)
        
    conn.commit()
    conn.close()

    return


# gets fullpath from source folder
def getxmlfiles(_path):
    
    xmlfilelist = []
    xmlsourcefolder = r'c:\ftp\sourcefolder'
    xmltargetfolder= r'c:\ftp\targetfolder'
    xmlfilename = ''

    xmlfilelist = [ xmlsourcefolder + os.sep + f for f in os.listdir(xmlsourcefolder) if f.endswith(r'.xml') or f.endswith(r'.XML')]
    #  get fullpath. 

    #  print(xmlfilelist)
    return xmlfilelist


# move impoted files into backup folder
def bakxmlfromlist(_xmlfilelist):
    
    xmltargetfolder= r'c:\ftp\targetfolder'
    xmlfilename = ''

    for i in _xmlfilelist:
        os.system('move ' + '"' + i + '"' + ' ' + xmltargetfolder)
    
    return 

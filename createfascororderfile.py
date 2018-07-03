import datetime
import shutil
import csv

class NextDetailLineInfo(object):
    """__init__() functions as the class constructor"""
    def __init__(self):
        self.NextLineNumber = 0
    
    def getNextLineNumber(self):
        self.NextLineNumber = self.NextLineNumber + 1
        return self.NextLineNumber

class FascorOrder2000(object):
    """__init__() functions as the class constructor"""
    def __init__(self):
        self.ControlID = ''     #1 is blank
        self.Type = 'O'         #1
        self.AccountNumber = '' #6
        self.DocNumber = ''     #6
        self.Date = ''          #6 yymmdd
        self.Seq = '0'          #1
        self.LineNo = '0'       #4
        self.Tran = '2000'      #4
        self.TimeStamp = get_CCYYMMDDHHMMSS()     #14  CCYYMMDDHHMMSS
        self.Facility = '01'    #2
        self.OrderId  = ''      # FMR-ORDER-ID-2000   PIC X(17).
        self.EmailAddress = ''  # FMR-EMAIL-ADDRESS-2000         PIC X(75).
        self.Filler = ''        # FILLER PIC X(888).

    def get2000(self):
         returnvalue = self.ControlID.ljust(1)[:1]\
           +self.Type.ljust(1)[:1]\
           +self.AccountNumber.ljust(6)[:6]\
           +self.DocNumber.ljust(6)[:6]\
           +self.Date.ljust(6)[:6]\
           +self.Seq.ljust(1)[:1]\
           +'{:0>4}'.format(str(self.LineNo))\
           +self.Tran.ljust(4)[:4]\
           +self.TimeStamp.ljust(14)[:14]\
           +self.Facility.ljust(2)[:2]\
           +self.OrderId.ljust(17)[:17]\
           +self.EmailAddress.ljust(75)[:75]\
           +self.Filler.ljust(888)[:888]
         return returnvalue

class FascorOrder1320(object):
    """__init__() functions as the class constructor"""
    def __init__(self):
        self.ControlID = ''     #1 is blank
        self.Type = 'O'         #1
        self.AccountNumber = '' #6
        self.DocNumber = ''     #6
        self.Date = ''          #6 yymmdd
        self.Seq = '0'          #1
        self.LineNo = '0'       #4
        self.Tran = '1320'      #4
        self.Mode = 'A'         #1
        self.Facility = '01'    #2
        self.OrderId = ''       #17
        self.DetailSeqNbr = self.Seq   # DETAIL-SEQ-NBR-1320             PIC X(09).
        self.Sku = ''                  # FMR-SKU-6-1320                  PIC X(15).
        self.SkuOrignalQty = 0         # FMR-SKU-ORIGINAL-QTY-1320       PIC X(08).
        self.SkuShipQty = 0            # FMR-SKU-SHIP-QTY-1320           PIC X(08).
        self.SkuUseReserveQty = 'N'    # FMR-SKU-USE-RESERVE-QTY-1320    PIC X(01).
        self.SkuClass = '01'           # FMR-SKU-CLASS-1320              PIC X(10).
        self.SkuUom = 'EA'             # FMR-SKU-UOM-1320                PIC X(08).
        self.SkuAllocationCntrl = 'L'  # FMR-SKU-ALLOCATION-CNTRL-1320   PIC X(01).
        self.PartialShipCntrl = 'Y'    # FMR-PARTIAL-SHIP-CNTRL-1320     PIC X(01).
        self.SkuFillPercentage = ''    # SKU-FILL-PERCENTAGE-1320        PIC X(08).
        self.SkuSubstituteCntrl = 'N'  # FMR-SKU-SUBSTITUTE-CNTRL-1320   PIC X(01).
        self.SkuSubstituteMix = 'N'    # FMR-SKU-SUBSTITUTE-MIX-1320     PIC X(01).
        self.FullCaseOnly = 'N'        # FMR-FULL-CASE-ONLY-1320         PIC X(01).
        self.SkuAbvDescription = ''    # FMR-SKU-ABV-DESCRIPTION-1320    PIC X(10).
        self.PrtLineOnly = 'N'         # FMR-PRT-LINE-ONLY-1320          PIC X(01).
        self.PrtComCntrlPlist = '0'    # FMR-PRT-COM-CNTRL-PLIST-1320    PIC X(01).
        self.PrtComCntrlBol = '0'      # FMR-PRT-COM-CNTRL-BOL-1320      PIC X(01).
        self.LineComment1 = ''         # FMR-LINE-COMMENT1-1320          PIC X(90).
        self.LineComment2 = ''         # FMR-LINE-COMMENT2-1320          PIC X(90).
        self.LineComment3 = ''         # FMR-LINE-COMMENT3-1320          PIC X(90).
        self.ClassOfGoodsMajor = ''    # FMR-CLASS-OF-GOODS-MAJOR-1320  PIC X(05).
        self.ClassOfGoodsMinor = ''    # FMR-CLASS-OF-GOODS-MINOR-1320  PIC X(05).
        self.SkuInsureValue = ''       # FMR-SKU-INSURE-VALUE-1320      PIC X(08).
        self.CustSkuId = ''            # FMR-CUST-SKU-ID-1320           PIC X(30).
        self.CustSkuDescription1 = ''  # FMR-CUST-SKU-DESCRIPTION1-1320 PIC X(40).
        self.CustSkuDescription2 = ''  # FMR-CUST-SKU-DESCRIPTION2-1320 PIC X(40).
        self.RetailUom = ''            # FMR-RETAIL-UOM-1320            PIC X(08).
        self.SkuCatalogPage = ''       # FMR-SKU-CATALOG-PAGE-1320      PIC X(12).
        self.SkuAssortmentNbr = ''     # FMR-SKU-ASSORTMENT-NBR-1320    PIC X(10).
        self.SkuRetailPrice = ''       # FMR-SKU-RETAIL-PRICE-1320      PIC X(08).
        self.SkuRetailUpcCd = ''       # FMR-SKU-RETAIL-UPC-CD-1320     PIC X(15).
        self.QuantityOfStickers = ''   # FMR-QUANTITY-OF-STICKERS-1320  PIC X(04).
        self.PriceStickerObject = ''   # FMR-PRICE-STICKER-OBJECT-1320  PIC X(08).
        self.LotMix = ''               # FMR-LOT-MIX-1320               PIC X(06).
        self.LotId = ''                # FMR-LOT-ID-1320                PIC X(15).
        self.LotMfgAge = ''            # FMR-LOT-MFG-AGE-1320           PIC X(06).
        self.LotMfgDate = ''           # FMR-LOT-MFG-DATE-1320          PIC X(08).
        self.LotMfgDateRange = ''      # FMR-LOT-MFG-DATE-RANGE-1320    PIC X(03).
        self.LotExpAge = ''            # FMR-LOT-EXP-AGE-1320           PIC X(06).
        self.LotExpDate = ''           # FMR-LOT-EXP-DATE-1320          PIC X(08).
        self.LotExpDateRange = ''      # FMR-LOT-EXP-DATE-RANGE-1320    PIC X(03).
        self.LotDistrRuleId = ''       # FMR-LOT-DISTR-RULE-ID-1320     PIC X(10).
        self.Filler = ''               # FILLER                         PIC X(363).


class FascorOrder1310(object):
    """__init__() functions as the class constructor"""
#   def __init__(self, text1=None, num1=None, text2=None):
    def __init__(self):
        self.ControlID = ' '    #1 is blank
        self.Type = 'O'         #1
        self.AccountNumber = '' #6
        self.DocNumber = ''     #6
        self.Date = ''          #6 yymmdd
        self.Seq = '0'          #1
        self.LineNo = '0'       #4
        self.Tran = '1310'      #4
        self.Mode = 'A'         #1
        self.Facility = '01'    #2
        self.OrderId = ''       #17
        self.DateCCYYMMDD = ''  #8
        self.OrderType = 'GPH'     #5
        self.Priority = ''      #2
        self.Status = 'OKAY'    #4
        self.ReasonCode = ''    #5
        self.Comments = ''      #30
        self.ShipDateCCYYMMDD = ''  #8  Needs by
        self.ShipDay = ''       #2  not used but must exist
        self.CarrierId = ''     #15
        self.ShippingTerms = 'PREPAID' #10
        self.ShipWogCntrl = 'N' #1 SHIP-WOG-CNTRL
        self.AsnCntrl = 'N'     #1
        self.RouteId = '15'       #15 ROUTE-ID
        self.StopId = ''        #5
        self.SoldToCustId = ''  #15 SOLD-TO-CUST-ID
        self.SoldToCustName = '' #30 SOLD-TO-CUST-NAME
        self.SoldToCustAbvName = '' #15 SOLD-TO-CUST-ABV-NAME-1310 PIC X(15)
        self.SoldToCustAddr1 = ''   #30 SOLD-TO-CUST-ADDR1
        self.SoldToCustAddr2 = ''   #30 SOLD-TO-CUST-ADDR1
        self.SoldToCustAddr3 = ''   #30 SOLD-TO-CUST-ADDR1
        self.SoldToCustAddr4 = ''   #30 SOLD-TO-CUST-ADDR1
        self.SoldToCustCity = ''
        self.SoldToCustState = ''
        self.SoldToCustZip5 = ''
        self.SoldToCustZip4 = ''
        self.SoldToCustCtyCode = ''
        self.SoldToCustPoRef = ''
        self.SoldToCustPoDate = ''
        self.ShipToCustId = ''  
        self.ShipToCustName = '' 
        self.ShipToCustAbvName = ''
        self.ShipToCustAddr1 = ''   
        self.ShipToCustAddr2 = ''   
        self.ShipToCustAddr3 = ''   
        self.ShipToCustAddr4 = ''   
        self.ShipToCustCity = ''
        self.ShipToCustState = ''
        self.ShipToCustZip5 = ''
        self.ShipToCustZip4 = ''
        self.ShipToCustCtyCode = ''
        self.ToteCntrl = 'N'
        self.ToteType = ''
        self.ChildType = ''
        self.PalletCntrl = 'N'
        self.PalletType = ''
        self.OrderFillCntrl = 'N'
        self.OrderFillPercentage = ''
        self.OrderFillCalc = ''
        self.OrderFillAction = ''
        self.PartialShipCntrl = 'Y' 
        self.PartialShipSuffix = '00' # PARTIAL-SHIP-SUFFIX-1310   PIC X(02).
        self.SubstituteCntrl = 'N'  #SUBSTITUTE-CNTRL-1310      PIC X(01).
        self.SubstituteSuffix = '00'  #SUBSTITUTE-SUFFIX-1310     PIC X(02).
        self.PrtPackListCntrl = 'Y' # FMR-PRT-PACKLIST-CNTRL-1310    PIC X(01). 
        self.PrtPackListObject = 'PP_MHCMU'  #PRT-PACKLIST-OBJECT-1310   PIC X(08).
        self.PrtShipLabelCntrl = 'Y' #PRT-SHIP-LABEL-CNTRL-1310  PIC X(01). 
        self.PrtShipLabelObject = ''  #PRT-SHIP-LABEL-OBJ-1310    PIC X(08).
        self.PrtPriceLabelCntrl = 'N'  #PRT-PRICE-LABEL-CNTRL-1310 PIC X(01).
        self.PrtPriceLabelObject = ''  #PRT-PRICE-LABEL-OBJ-1310   PIC X(08).
        self.PrtOtherDocCntrl = 'N'  #PRT-OTHER-DOC-CNTRL-1310   PIC X(01).
        self.PrtOtherDocObject = ''  #PRT-OTHER-DOC-OBJECT-1310  PIC X(08).
        self.CodInvoiceAmount = ''  #COD-INVOICE-AMOUNT-1310    PIC X(08). 
        self.DeclaredValue = ''      #DECLARED-VALUE-1310        PIC X(08).
        self.PostPickProc1Cntrl = 'N'  #POST-PICK-PROC1-CNTRL-1310 PIC X(01).
        self.PostPickProc1Pfile = ''  #POST-PICK-PROC1-PFILE-1310 PIC X(08).
        self.PostPickProc2Cntrl = 'N'  #POST-PICK-PROC2-CNTRL-1310 PIC X(01).
        self.PostPickProc2Pfile = ''  #POST-PICK-PROC2-PFILE-1310 PIC X(08).
        self.PostPickProc3Cntrl = 'N'  #POST-PICK-PROC3-CNTRL-1310 PIC X(01).
        self.PostPickProc3Pfile = ''  #POST-PICK-PROC3-PFILE-1310 PIC X(08).
        self.Filler = ''   # FILLER  PIC X(178).  Note: had to increase lenfth to 268 to match current line length.

        self.od1320 = []  # This holds the order detail records, FascorOrder1320
        
    def get1310(self):
         returnvalue = self.ControlID.ljust(1)[:1]\
           +self.Type.ljust(1)[:1]\
           +self.AccountNumber.ljust(6)[:6]\
           +self.DocNumber.ljust(6)[:6]\
           +self.Date.ljust(6)[:6]\
           +self.Seq.ljust(1)[:1]\
           +'{:0>4}'.format(str(self.LineNo))\
           +self.Tran.ljust(4)[:4]\
           +self.Mode.ljust(1)[:1]\
           +self.Facility.ljust(2)[:2]\
           +self.OrderId.ljust(17)[:17]\
           +self.DateCCYYMMDD.ljust(8)[:8]\
           +self.OrderType.ljust(5)[:5]\
           +self.Priority.ljust(2)[:2]\
           +self.Status.ljust(4)[:4]\
           +self.ReasonCode.ljust(5)[:5]\
           +self.Comments.ljust(30)[:30]\
           +self.ShipDateCCYYMMDD.ljust(8)[:8]\
           +self.ShipDay.ljust(2)[:2]\
           +self.CarrierId.ljust(15)[:15]\
           +self.ShippingTerms.ljust(10)[:10]\
           +self.ShipWogCntrl.ljust(1)[:1]\
           +self.AsnCntrl.ljust(1)[:1]\
           +self.RouteId.ljust(15)[:15]\
           +self.StopId.ljust(5)[:5]\
           +self.SoldToCustId.ljust(15)[:15]\
           +self.SoldToCustName.ljust(30)[:30]\
           +self.SoldToCustAbvName.ljust(15)[:15]\
           +self.SoldToCustAddr1.ljust(30)[:30]\
           +self.SoldToCustAddr2.ljust(30)[:30]\
           +self.SoldToCustAddr3.ljust(30)[:30]\
           +self.SoldToCustAddr4.ljust(30)[:30]\
           +self.SoldToCustCity.ljust(30)[:30]\
           +self.SoldToCustState.ljust(4)[:4]\
           +self.SoldToCustZip5.ljust(5)[:5]\
           +self.SoldToCustZip4.ljust(4)[:4]\
           +self.SoldToCustCtyCode.ljust(5)[:5]\
           +self.SoldToCustPoRef.ljust(17)[:17]\
           +self.SoldToCustPoDate.ljust(8)[:8]\
           +self.ShipToCustId.ljust(15)[:15]\
           +self.ShipToCustName.ljust(30)[:30]\
           +self.ShipToCustAbvName.ljust(15)[:15]\
           +self.ShipToCustAddr1.ljust(30)[:30]\
           +self.ShipToCustAddr2.ljust(30)[:30]\
           +self.ShipToCustAddr3.ljust(30)[:30]\
           +self.ShipToCustAddr4.ljust(30)[:30]\
           +self.ShipToCustCity.ljust(30)[:30]\
           +self.ShipToCustState.ljust(4)[:4]\
           +self.ShipToCustZip5.ljust(5)[:5]\
           +self.ShipToCustZip4.ljust(4)[:4]\
           +self.ShipToCustCtyCode.ljust(5)[:5]\
           +self.ToteCntrl.ljust(1)[:1]\
           +self.ToteType.ljust(5)[:5]\
           +self.ChildType.ljust(5)[:5]\
           +self.PalletCntrl.ljust(1)[:1]\
           +self.PalletType.ljust(5)[:5]\
           +self.OrderFillCntrl.ljust(1)[:1]\
           +self.OrderFillPercentage.ljust(8)[:8]\
           +self.OrderFillCalc.ljust(1)[:1]\
           +self.OrderFillAction.ljust(4)[:4]\
           +self.PartialShipCntrl.ljust(1)[:1]\
           +self.PartialShipSuffix.ljust(2)[:2]\
           +self.SubstituteCntrl.ljust(1)[:1]\
           +self.SubstituteSuffix.ljust(2)[:2]\
           +self.PrtPackListCntrl.ljust(1)[:1]\
           +self.PrtPackListObject.ljust(8)[:8]\
           +self.PrtShipLabelCntrl.ljust(1)[:1]\
           +self.PrtShipLabelObject.ljust(8)[:8]\
           +self.PrtPriceLabelCntrl.ljust(1)[:1]\
           +self.PrtPriceLabelObject.ljust(8)[:8]\
           +self.PrtOtherDocCntrl.ljust(1)[:1]\
           +self.PrtOtherDocObject.ljust(8)[:8]\
           +self.CodInvoiceAmount.ljust(8)[:8]\
           +'{:0>8}'.format(str(self.DeclaredValue))\
           +self.PostPickProc1Cntrl.ljust(1)[:1]\
           +self.PostPickProc1Pfile.ljust(8)[:8]\
           +self.PostPickProc2Cntrl.ljust(1)[:1]\
           +self.PostPickProc2Pfile.ljust(8)[:8]\
           +self.PostPickProc3Cntrl.ljust(1)[:1]\
           +self.PostPickProc3Pfile.ljust(8)[:8]\
           +self.Filler.ljust(268)[:268]
         return returnvalue

    def get1320(self):
        rtn = []
        for od in self.od1320:
            rtn.append(
                self.ControlID.ljust(1)[:1]\
                +self.Type.ljust(1)[:1]\
                +self.AccountNumber.ljust(6)[:6]\
                +self.DocNumber.ljust(6)[:6]\
                +self.Date.ljust(6)[:6]\
                +od.Seq.ljust(1)[:1]\
                +'{:0>4}'.format(str(od.LineNo))\
                +od.Tran.ljust(4)[:4]\
                +self.Mode.ljust(1)[:1]\
                +self.Facility.ljust(2)[:2]\
                +self.OrderId.ljust(17)[:17]\
               # use the value of LineNo. It should be the same as DetailSeqNbr
                +'{:0>9}'.format(str(od.LineNo))\
                +od.Sku.ljust(15)[:15]\
                +'{:0>8}'.format(str(od.SkuOrignalQty))\
               # use the value of SkuOrignalQty. It should be the same as SkuShipQty
                +'{:0>8}'.format(str(od.SkuOrignalQty))\
                +od.SkuUseReserveQty.ljust(1)[:1]\
                +od.SkuClass.ljust(10)[:10]\
                +od.SkuUom.ljust(8)[:8]\
                +od.SkuAllocationCntrl.ljust(1)[:1]\
                +od.PartialShipCntrl.ljust(1)[:1]\
                +od.SkuFillPercentage.ljust(8)[:8]\
                +od.SkuSubstituteCntrl.ljust(1)[:1]\
                +od.SkuSubstituteMix.ljust(1)[:1]\
                +od.FullCaseOnly.ljust(1)[:1]\
                +od.SkuAbvDescription.ljust(10)[:10]\
                +od.PrtLineOnly.ljust(1)[:1]\
                +od.PrtComCntrlPlist.ljust(1)[:1]\
                +od.PrtComCntrlBol.ljust(1)[:1]\
                +od.LineComment1.ljust(90)[:90]\
                +od.LineComment2.ljust(90)[:90]\
                +od.LineComment3.ljust(90)[:90]\
                +od.ClassOfGoodsMajor.ljust(5)[:5]\
                +od.ClassOfGoodsMinor.ljust(5)[:5]\
                +od.SkuInsureValue.ljust(8)[:8]\
                +od.CustSkuId.ljust(30)[:30]\
                +od.CustSkuDescription1.ljust(40)[:40]\
                +od.CustSkuDescription2.ljust(40)[:40]\
                +od.RetailUom.ljust(8)[:8]\
                +od.SkuCatalogPage.ljust(12)[:12]\
                +od.SkuAssortmentNbr.ljust(10)[:10]\
                +od.SkuRetailPrice.ljust(8)[:8]\
                +od.SkuRetailUpcCd.ljust(15)[:15]\
                +od.QuantityOfStickers.ljust(4)[:4]\
                +od.PriceStickerObject.ljust(8)[:8]\
                +od.LotMix.ljust(6)[:6]\
                +od.LotId.ljust(15)[:15]\
                +od.LotMfgAge.ljust(6)[:6]\
                +od.LotMfgDate.ljust(8)[:8]\
                +od.LotMfgDateRange.ljust(3)[:3]\
                +od.LotExpAge.ljust(6)[:6]\
                +od.LotExpDate.ljust(8)[:8]\
                +od.LotExpDateRange.ljust(3)[:3]\
                +od.LotDistrRuleId.ljust(10)[:10]\
                +od.Filler.ljust(363)[:363]\
            )
        return rtn

    def add_1320(self, object):
        self.od1320.append(object)

def get_YDDD():
    #dt = datetime.datetime(2006, 11, 21, 16, 30)
    dt = datetime.datetime.now()
    # Using datetime.timetuple() to get tuple of all attributes
    tt = dt.timetuple()
    yddd = str(tt.tm_year)[-1:] + '{:0>3}'.format(str(tt.tm_yday))
    return yddd

def get_YYDDD():
    # dt = datetime.datetime(2006, 11, 21, 16, 30)
    dt = datetime.datetime.now()
    # Using datetime.timetuple() to get tuple of all attributes
    tt = dt.timetuple()
    yyddd = str(tt.tm_year)[-2:] + '{:0>3}'.format(str(tt.tm_yday))
    return yyddd

def get_CCYYMMDDHHMMSS():
    # dt = datetime.datetime(2006, 11, 21, 16, 30)
    dt = datetime.datetime.now()
    # Using datetime.timetuple() to get tuple of all attributes
    tt = dt.timetuple()
    ccyymmddhhmmss = str(tt.tm_year) + '{:0>2}'.format(str(tt.tm_mon)) + '{:0>2}'.format(str(tt.tm_mday)) + '{:0>2}'.format(str(tt.tm_hour)) + '{:0>2}'.format(str(tt.tm_min)) + '{:0>2}'.format(str(tt.tm_sec))
    return ccyymmddhhmmss

def get_XDaysFromNowCCYYMMDD(adddays):
    dt = datetime.datetime.now() + datetime.timedelta(days=adddays)
    tt = dt.timetuple()
    ccyymmdd = str(tt.tm_year)[-4:] + '{:0>2}'.format(str(tt.tm_mon)) + '{:0>2}'.format(str(tt.tm_mday))
    return ccyymmdd

def CreateFileWatcherFileForFascor():
    outfile = get_OutPutFileName()
    f = open(outfile,"w")  
    
    # Write 1310 Header record
    f.write(oh.get1310() + '\n')  
    
    # Write 1320 Detail record(s)
    for od in oh.get1320():
        f.write(od + '\n')
    
    # Write 2000 Email record
    f.write(em.get2000() + '\n')
    
    f.close()
    
    # Rename the file extension from ___ to CVT so the file watcher will process it.
    shutil.move(outfile, outfile[0:32]+'.CVT')

def get_OutPutFileName():
# Create an output filename matching the format used by Fascor Order Watcher.  FASMSG_OES1_2018031617010667_SRC.___  Note:  ___ is replaced with CVT after write is complete.
    currentDT = datetime.datetime.now()
    return '{:12s}{:0>4}{:0>2}{:0>2}{:0>2}{:0>2}{:0>2}{:0>6}{:4s}'.format('FASMSG_ODOO_', 
                                                                          str(currentDT.year), 
                                                                          str(currentDT.month), 
                                                                          str(currentDT.day),
                                                                          str(currentDT.hour), 
                                                                          str(currentDT.minute), 
                                                                          str(currentDT.second), 
                                                                          str(currentDT.microsecond), 
                                                                          '.___' )

# Create 1310 Header Record
#TODO: Get from Odoo for production
oh = FascorOrder1310() 
oh.AccountNumber = '398860'
oh.DocNumber = 'S12345'
oh.Date = '180317' #yymmdd
#oh.OrderId = oh.AccountNumber + oh.DocNumber + get_YDDD() + 'O'
oh.OrderId = oh.AccountNumber + oh.DocNumber + get_YYDDD()
oh.DateCCYYMMDD = '20180317'
oh.Priority = 'PR'
oh.CarrierId = 'RUSH'  # leave blank to rate shop
if oh.CarrierId == 'RUSH':
   oh.ShipDateCCYYMMDD = get_XDaysFromNowCCYYMMDD(2)  # TODO: Tweak days based on level of service
else:
   oh.ShipDateCCYYMMDD = get_XDaysFromNowCCYYMMDD(30)  
oh.SoldToCustId = oh.AccountNumber #15 Normaly same as oh.AccountNumber 
oh.SoldToCustName = 'COMMUNITY CHURCH'
oh.SoldToCustAbvName = 'COMMUNITY CHURCH' #.ljust(15)[:15] #15  Note: SoldToCustAbvName.ljust(15)[:15] ensures a value is exactly 15 bytes long. 
oh.SoldToCustAddr1 = ''
oh.SoldToCustAddr2 = '808 REYNOLDS ST'
oh.SoldToCustAddr3 = ''
oh.SoldToCustAddr4 = ''
oh.SoldToCustCity = 'SPRINGHILL'
oh.SoldToCustState = 'LA'
oh.SoldToCustZip5 = '71075'
oh.SoldToCustZip4 = ''
oh.SoldToCustPoRef = 'LADYBUGS'
oh.ShipToCustName = 'COMMUNITY CHURCH NW' 
oh.ShipToCustAbvName = 'COMMUNITY CHURCH'
oh.ShipToCustAddr1 = '808 REYNOLDS ST'   
oh.ShipToCustAddr2 = ''   
oh.ShipToCustAddr3 = ''   
oh.ShipToCustAddr4 = ''   
oh.ShipToCustCity = 'SPRINGHILL'
oh.ShipToCustState = 'LA'
oh.ShipToCustZip5 = '71075'
oh.ShipToCustZip4 = '3548'
oh.DeclaredValue = 1299

# Create 1320 Detail Record(s)
detail_line = NextDetailLineInfo()  # Always initlize this before next order.
reader = csv.reader(open('OrderDetailsTest.csv', 'rt'))  #read csv file.  TODO: Get from Odoo for production
for row in reader:   # loop thru each record in file
    od = FascorOrder1320() 
    od.LineNo = detail_line.getNextLineNumber()
    od.Sku = row[0]
    od.SkuOrignalQty = row[1]
    od.SkuAbvDescription = row[2]
    oh.add_1320(od)

# Create 2000 Email Address Record
em = FascorOrder2000()
em.LineNo = detail_line.getNextLineNumber()
em.AccountNumber = oh.AccountNumber
em.DocNumber = oh.DocNumber
em.OrderId = oh.OrderId
em.Date = oh.Date
em.EmailAddress = 'downslj@gmail.com'

# Create an output filename matching the format used by Fascor Order Watcher.  FASMSG_ODOO_2018031617010667_SRC.CVT
CreateFileWatcherFileForFascor()  

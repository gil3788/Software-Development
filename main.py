import os
import collections
from Tkinter import *
from tkFileDialog import askopenfilenames
import tkMessageBox
from pdf_lib.pdf_py import pdfPageFunction, pdfValueFinder, pdfRepetitions
from excel_lib.excel import  excelFiscal, dcmMatchedExcel, saveExcel, dbmExcelSummary, reportOverview, dbmMatchedExcel , dcmExcelSummary
import yaml
#Global Dictonaries/Objects and Variables#
pdfDirectory = {}
dcmSummaryInvoices = {}
fiscalInvoices = {}
dbmMatchedInvoices = {}
dcmMatchedInvoices = {}
unmatchedInvoices = {}
campaignDataObject = {}
dbmSummaryInvoices = {}

FIND_OTN = "OTN"
FIND_SUMMARY_CLIENT_NUM ="Numero de Cliente"
FIND_SUMMARY = "Nro de Documento Interno"
FIND_FISCAL = "Folio fiscal"
FIND_FISCAL_TRANS = "Detalles de la transacci"
FIND_FISCAL_AMOUNT="TotalMXN$"
FIND_FISCAL_AMOUNT_END="eda transferencia bancaria"
FIND_AMOUNT = "Monto AdeudadoMXN"
FIND_SUMMARY_DFA = "DART for Advertisers"
FIND_SUMMARY_CPM = "CPM"
FIND_SUMMARY_IMP = "IMP"
FIND_SUMMARY_NOTES = "- Notes"
FIND_SUMMARY_ADVERTISER_NAME = "Anunciante"
FIND_SUMMARY_ADVERTISER_ID = "ID:"
FIND_SUMMARY_CAMPAIGN ='DoubleClick Campaign Manager'
FIND_SUMMARY_PRODUCT = "Producto"
FIND_SUMMARY_DATE = "Fecha:"
#############

def campaignDataExtracterDCM (tempPage, invoiceIndex) :
    ##added index to identify proper invoice##
    tempCampaignObj = {}
    #print "DCM Invoice Index: ", invoiceIndex
    index = "invoice"+str(invoiceIndex)
    counter = 0
    campaignList = [m.start() for m in re.finditer('DoubleClick Campaign Manager', tempPage["text"])]
    campaignDataObject ["numCampaigns"] = len(campaignList)
    while counter < len(campaignList) :
        campaignString = tempPage["text"][campaignList[counter]:len(tempPage["text"])]
        #print "campaign string: ",campaignString
        indexValueStart = campaignString.find("- Campa")
        indexValueEnd = campaignString.find("- Notas:")
        campaignName = campaignString[ indexValueStart : indexValueEnd]
        indexValueStart = campaignName.find("ID:")
        indexValueEnd = len(campaignName)
        campaignId = campaignName[ indexValueStart : indexValueEnd]
        campaignId = campaignId.replace("ID:", "")
        print "----------------------------------------------------Found CPC:---------------------------------------------------------- ", campaignString.find("CPC")

        indexValueStart = campaignString.find("- Notas:")
        indexValueEnd = campaignString.find("IMP")
        campaignImp = campaignString[ indexValueStart+11 : indexValueEnd]


        #print "Campaign Impressions: ",campaignImp

        tempCampaignObj ["campaign" + str(counter)]= {
            "campaignName" : campaignName[ 9: len(campaignName) ],
            "campaignId" : campaignId,
            "campaignImp": campaignImp,

        }
        campaignDataObject[index] = tempCampaignObj

        #print "Campaign from Data Extracter Function : "
        print yaml.dump(campaignDataObject, default_flow_style=False)
        counter = counter + 1

def campaignDataExtracterDBM (tempPage) :
    DataObject = {}
    counter = 0
    campaignList = [m.start() for m in re.finditer('Anunciante', tempPage["text"])]
    campaignDataObject ["numCampaigns"] = len(campaignList)
    while counter < len(campaignList) :
        campaignString = tempPage["text"][campaignList[counter]:len(tempPage["text"])]
        #print "campaignString ", campaignString

        indexValueStart = campaignString.find("- Campa")
        if campaignString.find("- Notas:") == -1:
            indexValueEnd = campaignString.find("1EA")
        else:
            indexValueEnd = campaignString.find("- Notas:")
        campaignName = campaignString[ indexValueStart : indexValueEnd]
        indexValueStart = campaignName.find("ID:")
        indexValueEnd = len(campaignName)
        campaignId = campaignName[ indexValueStart : indexValueEnd]
        campaignId = campaignId.replace("ID:", "")
        DataObject ["campaign" + str(counter)] = { "campaignName" : campaignName[ 10: len(campaignName) ], "campaignId" : campaignId }
        #print "Campaign from Data Extracter Function : ", campaignDataObject
        counter = counter + 1
    return DataObject

def getPdfFiles ():
    counter = 0
    for file in os.listdir("pdfAssets"):
        if file.endswith(".pdf"):
            pdfDirectory["pdf_file" + str(counter)] = os.path.join("pdfAssets", file)
            counter = counter + 1
    return pdfDirectory

def dbmInvoiceMatcher(summaryObject, fiscalObject) :
    summaryCounter = 0
    orderedCounter = 0
    while summaryCounter < len(summaryObject.keys()) :
        fiscalCounter = 0
        while fiscalCounter < len(fiscalObject.keys()) :
            if summaryObject["invoice"+str(summaryCounter)]["transactionId"] == fiscalObject["invoice"+str(fiscalCounter)] ["transactionId"]:

                dbmMatchedInvoices["invoice"+str(orderedCounter)] = {
                    "summaryInvoice" : summaryObject["invoice"+str(summaryCounter)],
                    "fiscalInvoice" : fiscalObject["invoice"+str(fiscalCounter)]
                }
                orderedCounter = orderedCounter + 1
            fiscalCounter = fiscalCounter + 1
        summaryCounter = summaryCounter + 1
    #print matchedInvoices

def dcmInvoiceMatcher(summaryObject, fiscalObject) :
    summaryCounter = 0
    orderedCounter = 0
    while summaryCounter < len(summaryObject.keys()) :
        fiscalCounter = 0
        while fiscalCounter < len(fiscalObject.keys()) :
            if summaryObject["invoice"+str(summaryCounter)]["transactionId"] == fiscalObject["invoice"+str(fiscalCounter)] ["transactionId"]:

                dcmMatchedInvoices["invoice"+str(orderedCounter)] = {
                    "summaryInvoice" : summaryObject["invoice"+str(summaryCounter)],
                    "fiscalInvoice" : fiscalObject["invoice"+str(fiscalCounter)]
                }
                orderedCounter = orderedCounter + 1
            fiscalCounter = fiscalCounter + 1
        summaryCounter = summaryCounter + 1
    #print matchedInvoices

def main() :
    pdfFilesNames = getPdfFiles()
    counter = 0
    dcmSummaryCounter = 0
    dbmSummaryCounter = 0
    fiscalCounter = 0

    while counter < len(pdfFilesNames.keys()) :
        pdfFileName = pdfFilesNames["pdf_file"+str(counter)]
        pdfFileName = pdfFileName.replace("pdfAssets/", "")
        tempPage = pdfPageFunction(pdfFilesNames["pdf_file"+str(counter)], 0)
        #print "temp page: ", tempPage
        valueFoundSummary = pdfValueFinder(tempPage["text"], FIND_SUMMARY)
        valueFoundFiscal = pdfValueFinder(tempPage["text"], FIND_FISCAL)
        valueFoundOtn = pdfValueFinder(tempPage["text"], FIND_OTN)

        if valueFoundSummary != -1:
            #print tempPage["text"]
            invoiceAmount = ""
            invoiceTax = ""
            product = ""
            date = ""
            advertiserName = ""
            advertiserId = ""
            keyId = tempPage["text"][valueFoundSummary+25:valueFoundSummary+34]
            valueFoundProduct = pdfValueFinder(tempPage["text"], FIND_SUMMARY_PRODUCT)
            valueFoundDART_DCM = pdfValueFinder(tempPage["text"], FIND_SUMMARY_DFA)

            #Determine if Invoice is from DCM Platform#
            if valueFoundDART_DCM != -1:
                valueFoundDate = pdfValueFinder(tempPage["text"], FIND_SUMMARY_DATE)
                valueBillingAmount = pdfValueFinder(tempPage["text"], FIND_AMOUNT)
                valueFoundClientId = pdfValueFinder(tempPage["text"], FIND_SUMMARY_CLIENT_NUM)
                valueFoundAdvertiser = pdfValueFinder(tempPage["text"], FIND_SUMMARY_ADVERTISER_NAME)
                valueFoundAdvertiserId = pdfValueFinder(tempPage["text"], FIND_SUMMARY_ADVERTISER_ID)
                valueFoundCampaign = pdfValueFinder(tempPage["text"], FIND_SUMMARY_CAMPAIGN)
                ##Call Campaign Data Extractor Function##
                campaignDataExtracterDCM(tempPage, dcmSummaryCounter)
                #########################
                valueFoundCPM = pdfValueFinder(tempPage["text"], FIND_SUMMARY_CPM)
                valueFoundIMP = pdfValueFinder(tempPage["text"], FIND_SUMMARY_IMP)
                billingAmount = tempPage["text"][valueBillingAmount:len(tempPage["text"])]
                indexEnd = billingAmount.find("Monto Pagado")
                invoiceAmount =  billingAmount [0:indexEnd].replace("Monto AdeudadoMXN", "")
                product = tempPage["text"][valueFoundProduct:valueFoundSummary].replace("Producto:" , "")
                product = product + " Campaign Manager"
                date = tempPage["text"][valueFoundDate:valueFoundClientId].replace("Fecha:" , "")
                advertiserString = tempPage["text"][valueFoundCampaign : len(tempPage["text"])]
                indexValueStart = advertiserString.find("Anunciante:")
                indexValueEnd = advertiserString.find("- Campa")
                advertiserName = advertiserString [ indexValueStart : indexValueEnd ].replace("Anunciante:", "")
                advertiserId = advertiserString [ advertiserString.find("ID:") : advertiserString.find("- Cam") ].replace("ID:", "")

                try:
                    invoiceTax = float(invoiceAmount.replace(",", ""))*.13793103#16% Mexican Tax#
                    invoiceTax = str(invoiceTax)
                except ValueError:
                    invoiceTax = "Error Float Operation or No Value Found"

                dcmSummaryInvoices["invoice"+str(dcmSummaryCounter)] = {
                        "transactionId" : keyId,
                        "invoiceAmount" : invoiceAmount,
                        "invoiceTax" : invoiceTax,
                        "product" : product,
                        "date" : date,
                        "advertiserName" : advertiserName,
                        "advertiserId" : advertiserId,
                        "campaignData" : campaignDataObject,
                        "fileName" : pdfFileName
                        #"campaignData" : campaignDataObject["invoice"+str(dcmSummaryCounter)],
                    }
                dcmSummaryCounter = dcmSummaryCounter + 1
            #Determine if Invoice is from DBM Platform#
            if valueFoundDART_DCM == -1:
                valueFoundDate = pdfValueFinder(tempPage["text"], FIND_SUMMARY_DATE)
                valueBillingAmount = pdfValueFinder(tempPage["text"], FIND_AMOUNT)
                valueFoundClientId = pdfValueFinder(tempPage["text"], FIND_SUMMARY_CLIENT_NUM)
                valueFoundAdvertiser = pdfValueFinder(tempPage["text"], FIND_SUMMARY_ADVERTISER_NAME)
                valueFoundAdvertiserId = pdfValueFinder(tempPage["text"], FIND_SUMMARY_ADVERTISER_ID)

                ##Call Campaign Data Extractor Function##
                dbmCampaigns = campaignDataExtracterDBM (tempPage)
                #print dbmCampaigns.keys()
                #########################
                #valueFoundCPM = pdfValueFinder(tempPage["text"], FIND_SUMMARY_CPM)
                #valueFoundIMP = pdfValueFinder(tempPage["text"], FIND_SUMMARY_IMP)
                billingAmount = tempPage["text"][valueBillingAmount:len(tempPage["text"])]
                indexEnd = billingAmount.find("Monto Pagado")
                invoiceAmount =  billingAmount [0:indexEnd].replace("Monto AdeudadoMXN", "")
                product = tempPage["text"][valueFoundProduct:valueFoundSummary].replace("Producto:" , "")
                date = tempPage["text"][valueFoundDate:valueFoundClientId].replace("Fecha:" , "")
                advertiserString = tempPage["text"][valueFoundAdvertiser : len(tempPage["text"])]
                #print "Debugging Console DBM Invoice Section: ", advertiserString
                indexValueStart = advertiserString.find("Anunciante:")
                indexValueEnd = advertiserString.find("- Campa")
                advertiserName = advertiserString [ indexValueStart : indexValueEnd ].replace("Anunciante:", "")
                advertiserId = advertiserString [ advertiserString.find("ID:") : advertiserString.find("- Cam") ].replace("ID:", "")
                if dbmCampaigns.keys() == 0 :
                    advertiserName ="Rebilled Invoice"
                    advertiserId = "Rebilled Invoice"
                try:
                    invoiceTax = float(invoiceAmount.replace(",", ""))*.13793103#16% Mexican Tax#
                    invoiceTax = str(invoiceTax)
                except ValueError:
                    invoiceTax = "Error Float Operation or No Value Found"

                dbmSummaryInvoices["invoice"+str(dbmSummaryCounter)] = {
                        "transactionId" : keyId,
                        "invoiceAmount" : invoiceAmount,
                        "invoiceTax" : invoiceTax,
                        "product" : product,
                        "date" : date,
                        "advertiserName" : advertiserName,
                        "advertiserId" : advertiserId,
                        "campaignData" : dbmCampaigns,
                        "fileName" : pdfFileName
                    }
                dbmSummaryCounter = dbmSummaryCounter + 1



        if valueFoundFiscal != -1 or valueFoundOtn != -1:
            #print tempPage["text"]
            valueFoundTrans = pdfValueFinder(tempPage["text"], FIND_FISCAL_TRANS)
            #print valueFoundTrans
            keyId = tempPage["text"][valueFoundTrans +26 : valueFoundTrans+35]
            otn = tempPage["text"][valueFoundOtn+5 : valueFoundOtn+14]

            if valueFoundFiscal != -1:
                #print "Fiscal Transaction Id: ", keyId
                valueBillingAmount= pdfValueFinder(tempPage["text"], FIND_FISCAL_AMOUNT)
                billingAmount = tempPage["text"][valueBillingAmount : valueBillingAmount+50]
                facturaString = tempPage["text"][0 : len(tempPage["text"])]
                #print facturaString
                indexValueStart = facturaString.find("FACTURA")
                indexValueEnd = facturaString.find("Fecha y hora de emisi")
                #print indexValueStart
                #print indexValueEnd
                facturaId = facturaString[indexValueStart+7:indexValueEnd]
                #print facturaId
                indexA = billingAmount.find("transferencia")
                #print indexA
                invoiceAmount =  billingAmount[9:indexA-4]
                try:
                    invoiceTax = float(invoiceAmount.replace(",", ""))*.1379#16% Mexican Tax#
                    invoiceTax = str(invoiceTax)
                except ValueError:
                    invoiceTax = "Error Float Operation: No Value Found"

                fiscalInvoices["invoice"+str(fiscalCounter)] = {
                "transactionId" : keyId,
                "invoiceTotal" : invoiceAmount,
                "invoiceTax" : invoiceTax,
                "facturaId": facturaId ,
                "fileName" : pdfFileName
                }
            else:
                #print "Fiscal OTN: ", otn
                invoiceTotal = float(invoiceAmount.replace(",", ""))*.1379#16% Mexican Tax#
                invoiceTotal = str(invoiceTotal)
                fiscalInvoices["invoice"+str(fiscalCounter)] = {
                "transactionId" : otn,
                "fileName" : pdfFileName
                }

            fiscalCounter = fiscalCounter + 1

        counter = counter + 1

    reportOverview()
    print "----------------------------------------------------------------"
    print yaml.dump(dcmSummaryInvoices, default_flow_style=False)
    #print "DCM Summary Invoices: ", dcmSummaryInvoices
    dcmExcelSummary(dcmSummaryInvoices)
    print "----------------------------------------------------------------"

    print "----------------------------------------------------------------"
    #print "DBM Summary Invoices: ", dbmSummaryInvoices
    #print yaml.dump(dbmSummaryInvoices, default_flow_style=False)
    #print dbmSummaryInvoices['invoice0']['campaignData']['campaign0']['campaignId']
    dbmExcelSummary(dbmSummaryInvoices)
    print "----------------------------------------------------------------"

    print "----------------------------------------------------------------"
    #print "Fiscal Invoices: ", fiscalInvoices
    #print yaml.dump(fiscalInvoices, default_flow_style=False)
    excelFiscal(fiscalInvoices)
    print "----------------------------------------------------------------"
    dbmInvoiceMatcher(dbmSummaryInvoices, fiscalInvoices)
    #dcmInvoiceMatcher(dcmSummaryInvoices, fiscalInvoices)
    dbmMatchedExcel(dbmMatchedInvoices)
    dcmMatchedExcel(dcmMatchedInvoices)
    saveExcel ("Billing Report")
    print "----------------------------------END PROGRAM---------------------------------------"

def fileUpload():
    while True:
        uploadedfilenames = askopenfilenames(multiple=True)
        if uploadedfilenames == '':
            tkMessageBox.showinfo(message="File Upload has been cancelled. Program will retun to Menu")
            return
        uploadedfiles = window.splitlist(uploadedfilenames)
        print uploadedfiles
        return uploadedfiles

def exitSolution() :
    quit()
#create window#
window = Tk()
#modify window#
window.title("Billing Solution")
window.geometry("300x250")

app = Frame(window).grid()
title = Label(app, text="Latam Platforms Billing Solution").grid()
runSolutionButton = Button(app, text="Run Solution",command=lambda:main()).grid()
#uploadButton = Button(app, text="Upload PDF Files",command=lambda:fileUpload()).grid()
exitButton = Button(app, text="Exit Solution",command=lambda:exitSolution()).grid()
#Event Loop Window GUI#
window.mainloop()
#main()

#Required Modules to get Credentials#
import httplib2
from httplib2 import Http
from apiclient.discovery import build
from oauth2client.service_account import ServiceAccountCredentials
import json
import argparse
from googleapiclient import discovery
from oauth2client import client
from oauth2client import file as oauthFile
from oauth2client import tools
from pdf_py import pdfPageFunction, pdfValueFinder

def dcm_get_report_file(service, report_id, file_id):
    request = service.files().get_media(reportId=report_id, fileId=file_id)
    response = request.execute()
    return response

def dcm_get_report_files_list(service, profile_id, report_id):
    request = service.reports().files().list(profileId=profile_id, reportId=report_id)
    response = request.execute()
    counter = 0
    for report_file in response ['items']:
        if report_file['status'] == 'REPORT_AVAILABLE'  and counter < 10:
            print ('Report file with ID %s and file name "%s" has status %s.' % (report_file['id'], report_file['fileName'], report_file['status']))
            counter = counter + 1
        else:
            break

    return response['items']

def dcm_get_report(service, profile_id, report_id):
    request = service.reports().get(profileId=profile_id, reportId=report_id)
    response = request.execute()
    print "Report Obtained By API: ", response
    return response

def dcm_run_report(service, profile_id, report_id):
    request = service.reports().run(profileId=profile_id, reportId=report_id)
    response = request.execute()
    print "Response from API: ", response
    return response

def dcm_report_creator (service, profile_id, account_id, report_name):
    report = {
      'accountId' : account_id,
       'name': report_name,
        'type': 'STANDARD',
        'criteria': {
            'dateRange': {'relativeDateRange': 'YESTERDAY'},
            'dimensions': [{'name': 'dfa:campaign'}],
            'metricNames': ['dfa:clicks']
                }
            }

    request = service.reports().insert(profileId=profile_id, body=report)
    # Execute request and print response.
    response = request.execute()
    return response

def dcm_api_service_creator () :

    scopes = ['https://www.googleapis.com/auth/dfareporting']

    #Get credentials using service_account json and store them#
    credentials = ServiceAccountCredentials.from_json_keyfile_name('service-account.json', scopes)
    http_auth = credentials.authorize(Http())
    ##troubleshooting print##
    #print credentials
    #print http_auth
    dcm_service = build('dfareporting', 'v2.7', http=http_auth)
    return dcm_service


dcm_service = dcm_api_service_creator();
profileId = 2955404
fileId = 540228987
reportId = 85483310
accountId ="7480"
reportName = "New Standard Report"
#response = dcm_report_creator(dcm_service, profileId, accountId, reportName);
#report_response = dcm_run_report(dcm_service, profileId, response["id"])
##print "report_response : ", report_response
dcm_get_report(dcm_service, profileId, reportId)
dcm_get_report_files_list(dcm_service, profileId, reportId)
report_file = dcm_get_report_file(dcm_service, reportId, fileId)
print "Response Report File: ", report_file

indexA = report_file.find("Grand Total")
indexB = report_file.find(":", indexA+1, len(report_file))
print report_file[indexB: len(report_file)]

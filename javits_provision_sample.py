from dnacentersdk import api

import openpyxl
import pprint
import getpass
import requests, requests.packages
from requests.packages.urllib3.exceptions import InsecureRequestWarning
requests.urllib3.disable_warnings()
import time

#dnac_ip = input("Enter DNAC IP \n")
dnac_ip = "10.29.203.251"
#dnac_admin = input("Enter DNAC Admin \n")
dnac_admin = 'admin'
'''try:
    password = getpass.getpass()
except Exception as error:
    print("Error ", error)'''
password = "Google1234@"

url = "https://"+dnac_ip+":443"
dnac = api.DNACenterAPI(base_url="https://"+dnac_ip+":443",version = "2.1.2",
username=dnac_admin,password=password,verify=False)
rows = []

def read_input_file():
    print("reading input file")
    wb_obj = openpyxl.load_workbook(filename="host_onboarding.xlsx")
    sheet_obj = wb_obj.active
    for i in sheet_obj.rows:
        rows.append(i)

def host_onboard_port(switch,port):
    for row in rows[1:]:
        if (row[0].value == switch) and row[2].value == port:
            device_management_ip = row[1].value
            data_pool = row[3].value
            voice_pool = row[4].value
            auth_template = row[5].value
    device_info = dnac.sda.get_device_info(device_management_ip)
    pprint.pprint(device_info.response)
    print("Provisioning Port")
#the below is a 
    payload =  [{
        "siteNameHierarchy":  device_info.response["siteNameHierarchy"],
        "deviceManagementIpAddress": device_management_ip,
        "interfaceName": port,
        "dataIpAddressPoolName": data_pool,
        "voiceIpAddressPoolName": voice_pool,
        "authenticateTemplateName": auth_template
        }]
    pprint.pprint(payload)
    resp = dnac.sda.add_port_assignment_for_user_device(payload = payload)
    print(resp)
    print(resp.response)
#########





def host_onboard_switch(switch):
    payload = []
    for row in rows[1:]:
        if (row[0].value == switch) :
            device_management_ip = row[1].value
            break
    device_info = dnac.sda.get_device_info(device_management_ip)
    for row in rows[1:]:
        if (row[0].value == switch) and (row[6].value == 'Yes'):
            payload = [{
                "siteNameHierarchy": device_info.response["siteNameHierarchy"],
                "deviceManagementIpAddress": row[1].value,
                "interfaceName":row[2].value,
                "dataIpAddressPoolName": row[3].value,
                "voiceIpAddressPoolName": row[4].value,
                "authenticateTemplateName": row[5].value
            }]
            print("provisioning below information")
            pprint.pprint(payload)
            time.sleep(1)
            resp = dnac.sda.add_port_assignment_for_user_device(payload=payload)
            print(resp)


def host_onboard():
    switches = []
    # switch = {'hostname1':{'ip':'10.3.40.5',site_hierarchy = 'Global/Amer/SJC'},'hostname2':{'ip':'10.3.40.6',site_hierarchy = 'Global/Amer/SJC'} }
    switches.append(rows[1][0].value)
    print(switches)
    for row in rows[1:]:
        flag = 0
        if row[0].value in switches:
            flag = 1
        if flag == 0:
            switches.append(row[0].value)

    print(switches)
    for switch in switches:
        print("starting host onboarding for switch ", switch)
        host_onboard_switch(switch)



if __name__ == '__main__':
    read_input_file()
    #host_onboard_switch("m-e-c9348-02")
    host_onboard()
    #host_onboard_port("m-e-c9348-01","TwoGigabitEthernet1/0/13")

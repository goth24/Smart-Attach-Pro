__author__ = 'ZA028309'


server= "qualitycenter.cerner.com/qcbin"
username= "za028309"
password= "za242902@C"
domainname= "IP"
projectname= "TD_VALIDATION_TESTS"


import qcapi, getpass
bug = qcapi.QC()
bug.username = username
bug.password = password
bug.qc_server = server
bug.qc_domainname = domainname
bug.qc_projectname = projectname
result = bug.query(columns=['BG_SUMMARY','BG_STATUS'], search={'BG_RESPONSIBLE': 'user', 'BG_STATUS': 'new OR open'})
for id in result.keys():
    print('%s: %s' % (id, result[id]['BG_SUMMARY']))
    print('Status: %s' % (result[id]['BG_STATUS']))

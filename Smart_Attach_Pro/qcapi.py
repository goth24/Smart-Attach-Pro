#!/bin/env python

'''
qcapi.py - Callable API-like interface to Quality Center 10.

This attempts to facilitate a pure Python scriptable interface for Quality Center 10.  This work is not based on any documentation on QC, but rather by simply reverse engineering the application.

This works for what I use it for. Some field customization may be required.  Basic functionality exists, but not everything.  YMMV.  Please provide input and suggestions.

Quick howto:

import qcapi
bug = qcapi.QC(id=12345,username='foo',password='bar')
bug.set('field1name=value')
bug.commit()
 -> connect
 -> send data
 -> disco

'''

import gzip
import hashlib
try:
  import http.client as httplib
except:
  import httplib
import logging
import os
import re
import sys
import time

__author__ = 'Jason Avery <jason.avery@gmail.com>'
__date__ = '2011-06-12'
__version__ = '0.90'
__license__ = {'type': 'GNU GPL v3', 'owner': 'Jason Avery', 'year': '2012', 'url': 'http://www.gnu.org/copyleft/gpl.html'}
__source__ = 'http://code.google.com/p/py-qcapi/'


class QC():
  '''QC(bugid=0,username,password)

Call this class like:

bug = qcapi.QC(id=12345) # (or id=0 for new)
bug.username = 'myusername'
bug.password = 'abc123'
bug.qc_server = 'qc.example.com'
bug.qc_domainname = 'FOO'
bug.qc_projectname = 'BAR'
# The fields are likely dependant on your QC setup.
bug.set({'field1': 'data1', 'field2': 'data2'})
bug.commit()
# return the new ticket number
print bug.id

'''
  def __init__ (self, id=0, username='', password='', log='info', cert=''):
    #threading.Thread.__init__(self)
# Feel free to tweak these variables within *your* script as needed.  Typically, the defaults are fine.
    self.id = int(id)
    self.username = username
    self.password = password
    self.cert = cert
    self.qc_baseurl = '/qcbin'
    self.use_ssl = False
    self.user_agent = 'TeamSoft WinInet Component'

    # You need to set these in your script.
    self.qc_server = ''
    self.qc_domainname = ''
    self.qc_projectname = ''
    # Sets the default case owner when creating new cases
    self.default_case_owner = ''
    # Sets the defaults for new cases.  This can be tweaked by using the set() function.
    self.default_new_case_fields = {
      'BG_ACTUAL_FIX_TIME': '0',
      'BG_USER_TEMPLATE_02': 'N',
      'BG_USER_TEMPLATE_11': re.sub('\.', '_', re.sub('_[^_]+$', '', self.username)), # SCM Submitter
      'BG_RESPONSIBLE': self.default_case_owner,
      'BG_USER_TEMPLATE_12': self.default_case_owner,
      'BG_USER_TEMPLATE_13': '',  #CC list
      'BG_DETECTION_DATE': self._current_date(),
      'BG_STATUS': 'New',
      'BG_USER_01': 'Customer',
      'BG_USER_02': 'Not Applicable',
      'BG_USER_03': 'New',
      'BG_USER_05': 'Support - Normal Usage',
      'BG_USER_13': 'None',
      'BG_USER_16': 'Task',
      'BG_DETECTED_BY': self.username,
      'BG_REPRODUCIBLE': 'Y',
      'BG_DESCRIPTION': '<html><body><b>Insert case information here.</b></body></html>',
      'BG_SUMMARY': 'Untitled',  # Case Title
      'BG_SEVERITY': '3 - Medium'
    }

    # See logging's docs for full logging level options
    # logging.ERROR = error messages only
    # logging.WARN = warning and error messages only
    # logging.INFO = informational, warning, error messages. Default here.
    # logging.DEBUG = everything.  calling QC(debug=True) does the same thing.
    # logging.NONE = No logging output.
    self.verbosity = logging.INFO
    if log == 'info':
      self.verbosity = logging.INFO
    elif log == 'warn':
      self.verbosity = logging.WARN
    elif log == 'error':
      self.verbosity = logging.ERROR
    elif log == 'debug':
      self.verbosity = logging.DEBUG
    elif log == 'none':
      # technically there is no logging.NONE. A value > 51 does the trick.
      self.verbosity = 100
    # set True for debugging breakpoints
    self._bp = False

# Please do not modify these variables.  There's really no need to.
    self.qc_connected = 0
    self.commandcount = 1
    self.name = str(id)
    self.data = {}
    self.olddata = {}
    self.cookie = ''
    self.hostname = 'INGLIP_SUMMONED' # IT HAS BEGUN
    self.qc_port = 80
    if self.use_ssl:
      self.qc_port = 443
    self.sessionkey = ''
    self.loginsessionid = '-1'
    self.projectsessionid = '-1'
    self._debug = '' # debugging use

    # init logging
    logging.basicConfig(datefmt='%H:%M:%S',format='[%(levelname)s %(asctime)s] %(message)s', level=self.verbosity)


  def breakpoint(self, bp=False):
    if bp:
      try:
        input('breakpoint:')
      except KeyboardInterrupt:
        raise
    return


#  def commit(self,sendmail=False):
  def commit(self):
    '''commit()

This commits any changes in self.data. Opens a case if no self.id is set.
'''
    try:
      if not self.data.keys():
        logging.error('You need to set some values for this. See qcapi.set.__doc__')
    except:
      logging.error("You accidentally self.data. You need to set some things to change with set()")

    payload=''

    loggedin = False
    if not self.sessionkey:
      try:
        self.http_connect(self.username,self.password)
      except Exception as e:
        logging.error("failed to login: %s" % e)
    else:
      loggedin = True

# the rest of this should be handled by other subfunctions, but I'm pressed for time right now

    # The description is meant to be in HTML tags. If it wasn't by the user (which is expected), fix it.
    try:
      if not re.match('<html>', self.data['BG_DESCRIPTION']):
        self.data['BG_DESCRIPTION'] = '<html><body><b>%s</b></body></html>' % self.data['BG_DESCRIPTION']
    except KeyError:
      pass

    oldvalues = {}
    oldcomments = ''
    try:
      if self.id is not 0:
        oldvalues = self.query(self.id, self.data.keys())
    except:
      pass
   
    # very noisy, enable if needed
    #logging.debug('oldvalues: %s' % oldvalues)
    #self._oldvalues = oldvalues

    try:
      self.data['BG_DEV_COMMENTS']
      try:
        if not re.match('<font', self.data['BG_DEV_COMMENTS']):
          curdate = self._current_date()
  # FIXME I don't know the user's full name for the comment.  Oh well..
          self.data['BG_DEV_COMMENTS'] = '<font color="#000080"><b>&lt;%s&gt;, %s:</b></font> %s' % (self.username, curdate, self.data['BG_DEV_COMMENTS'])
        if oldvalues[str(self.id)]['BG_DEV_COMMENTS']:
          oldcomments = re.findall('(?s)^<html><body>(.+)</body></html>$', oldvalues[str(self.id)]['BG_DEV_COMMENTS'])[0]
      except KeyError:
        pass
  
  # TODO need to merge certain fields together here:  BG_DEV_COMMENTS
  # TODO this should be moved into _buildpayload('PostBug')
  # TODO this merging should be replaced with a helper to sort out the proper html formatting
        try:
          #logging.debug('oldcomments: %s' % oldcomments)
  # FIXME this may be bugged and erroneously add extra <html><body> tags ... but looks ok in the gui.
          self.data['BG_DEV_COMMENTS'] = '<html><body>%s<br><font color="#000080"><b>%s</b></font><br>%s</body></html>' % (oldcomments, '_' * 40, self.data['BG_DEV_COMMENTS'])
          #logging.debug('newcomments: %s' % newcomments)
        except KeyError as e:
          logging.debug("Error parsing oldvalues: %s" % e)
      except Exception as e: 
        logging.error("failed to update bug: %s" % e)
      logging.debug('!!! data is now: %s' % self.data)
      self.breakpoint(self._bp)
    except KeyError:
      pass

    self._http_call(params=self._buildpayload('ObjectLock', data={'id': self.id, 'fields': self.data.keys()}))
    result = self._http_call(params=self._buildpayload('PostBug', data={'id': self.id, 'fields': self.data, 'oldvalues': oldvalues}))
    if not re.match('\\x5c[0-9A-Fa-f]{8}\\x5c14:str:{\r\nID:', result[0]):
      logging.error('Failed to PostBug! Error: %s' % result[3])
      raise
    if self.id == 0:
      # save new case number
      self.id = int(re.findall('(?s)^\\x5c[0-9A-Fa-f]{8}\\x5c14:str:{\r\nID:([0-9]+)', result[0])[0])
    self._http_call(params=self._buildpayload('ObjectUnlock', data={'id': self.id}))

    if sendmail:
      pass
  # FIXME This actually isn't working... MailEntity command isn't returning a valid response from the server, so I think its broke.  Maybe.
      '''
      try:
        owner = {}
        owner = self.query(self.id, ['BG_USER_TEMPLATE_12','BG_USER_TEMPLATE_13'])
        logging.debug('owner: %s' % owner)
        try:
          mail = {}
  # FIXME cclist is broke for right now: 'cclist': self.data['BG_USER_TEMPLATE_11'], 
          mail = {'mailto': owner[str(self.id)]['BG_USER_TEMPLATE_12'], 'cclist': owner[str(self.id)]['BG_USER_TEMPLATE_13'], 'id': self.id}
          result = self._http_call(params=self._buildpayload('MailEntity', data=mail))
        except KeyError:
          logging.warn("Can't send email, no owner defined!") # is this what keyerror means?  its in owner i think
      except Exception as e:
        logging.warn("failed at MailEntity: %s" % e)
      '''

    if not loggedin:
      try:
        result = self.http_disconnect()
      except Exception as e:
        logging.error("failed to logout: %s" % e)

    return None


  def query(self, id=0, columns=[], search={}, limit=0):
    '''query(12345,['Col1', 'Col2', 'Col3'], {'this': 'that')

Queries QC for the given array of column names for the case with the given ID.  Returns a dictionary of 'column': 'value' pairs.
'''
    loggedin = False
    if not self.sessionkey:
      self.http_connect(self.username,self.password)
    else:
      loggedin = True

    # do id, column scrubbing here

    result = []
    result = self._http_call(params=self._buildpayload('GetBugValue', data={'id': id, 'fields': columns, 'search': search, 'limit': limit}))
    
    if re.match('\\x221:str:0\\x22', result[0]):
      logging.error('Failed to perform query (server error).')
      logging.debug('Server error was: %s' % result[3])
      self.breakpoint(self._bp)

    raw = []
    raw2 = []
    raw3 = []
    output = {}
    # FIXME This method of splitting data may be a problem if a field actually has a [,:]. May need to add a (?<!\\) in the splits.

    raw = re.findall('(?s){\r\nFIELDS:\\x5c[0-9A-Fa-f]{8}\\x5c{(.+?)}(?:,\r\n|$)', result[0])

    for x in raw:
      #logging.debug('Raw: %s' % x)
      #self.breakpoint(self._bp)
      raw2.append(re.sub('\\x5c[0-9A-Fa-f]{8}\\x5c', '', x))

    for x in raw2:
      #logging.debug('Raw2: %s' % x)
      #self.breakpoint(self._bp)
      raw3.append(re.findall('(?s)(?:^|,)\r\n([^,:]+:(?:\\x22[^\\x22]*\\x22|.*?))(?=,?\r\n)', x))

    for x in raw3:
      #logging.debug('Raw3: %s' % x)
      #self.breakpoint(self._bp)
      for y in x:
        #logging.debug('DEBUG: %s ;;; %s' % (x,y))
        key,value = re.findall('(?s)^([^:]+):(.*)$', y)[0]
        if key == 'BG_BUG_ID':
          output[value] = {}
          for y in x:
            key2,value2 = re.findall('(?s)^([^:]+):(.*)$', y)[0]
            output[value][key2] = value2
          continue

    if not loggedin:
      try:
        self.http_disconnect()
      except Exception as e:
        logging.warn("failed to logout: %s" % e)

    return output


  def set(self,input):
    '''set(input)
    
Sets the fields and values to change before using commit().  You can give a string 'field1=data1' or a dictionary {'field1': 'data1'}.

Example:
bug = QC()
bug.set({'field1': 'data1', 'field2': 'data2'})
bug.set('field3=data3')
bug.commit()

'''
    if isinstance(input, dict):
      self.data = dict(self.data.items() + input.items())
    elif isinstance(input, str):
      key, value = re.findall('(?s)^([^=]+)=(.+)$', input)[0]
      self.data[key] = value
    else:
      logging.error('Unknown input type. Try: set("field=value") or set({"field1": "value1", "field2": "value2"})')
    return


  def qclen(self,string,upper=False):
    '''Returns the length value of a given string.  The length is used because the QC devs don't understand string terminatation or character escaping.

length = self.qclen(string)
'''
    length1 = '0' * 8
    length = re.sub('^0x', '', hex(len(str(string))))
    length = length1[len(length)-8:] + length
    if upper:
      length = length.upper()
    return length


  def _current_date(self):
    return time.strftime('%Y-%m-%d')


  def _buildpayload(self,command,data={}):
    '''Builds the data format before POST'ing to QC. Not really meant to be called directly.  There's probably a better way to marshall the data, but I'm lazy and not feeling too fancy.  Possibly a bit of gas.  Anyways, set data key,value to what you want to send.

In short, you're not required know how this function really works.

Valid commands:
Login
ConnectProject: Part of login process
GetCommonSetting (schema query)
GetProjectCustomizationData:  Returns all fields, data types, pull down options, easily a metric shit tonnes of data.   Sometimes it just acts like a ping every 4 seconds (!!!), I guess the local app is requesting for schema updates. A lot.
GetBugValue: Get data
ObjectLock: locks fields, need to do before PostBug
ObjectUnlock: opposite of lock
PostBug: Push data
MailEntity: Free spam!
DisconnectProject: Part of logout process
Logout: gtfo

GetTreeNodeValue: gets pulldown menu info for the GUI
PingtoServer: does nothing
'''
    payload = '{\r\n'
    if command is 'Login':
      # data={'username': username, 'password': password} # this should come from https_connect()
      try:
        data['username']
        data['password']
        if not isinstance(data['username'], str):
          raise KeyError
        if not isinstance(data['password'], str):
          raise KeyError
      except KeyError as e:
        logging.error('Missing or incorrect input for QC._buildpayload(\'%s\').  See QC._buildpayload.__doc__ : %s' % (command,e))
      except:
        logging.error('General fail in QC._buildpayload(%s).' % command)
      payload += '0: \\%s\\%s,\r\n' % (self.qclen('0:conststr:%s' % command), '0:conststr:%s' % command)
      payload += '1: "0:int:%s",\r\n' % self.commandcount      # number of commands issued?
      payload += '2: "0:int:%s",\r\n' % self.loginsessionid
      payload += '3: "0:conststr:%s",\r\n' % self.sessionkey
      payload += '4: "0:int:%s",\r\n' % self.projectsessionid
      creds = '0:conststr:{\r\n'
      creds += 'USER_NAME:\\%s\\%s,\r\n' % (self.qclen(data['username']),data['username'])
      encpasswd = self.password_encrypt(data['password'])
      creds += 'PASSWORD:\\%s\\%s,\r\n' % (self.qclen(encpasswd), encpasswd)
      creds += 'CLIENTTYPE:OTAClient\r\n}\r\n'
      payload += '5: \\%s\\%s,\r\n' % (self.qclen(creds, upper=True), creds)
      hostname = '0:conststr:%s' % self.hostname
      payload += '6: \\%s\\%s,\r\n' % (self.qclen(hostname), hostname)
      payload += '7: "65536:str:0",\r\n8: "0:pint:0",\r\n9: "65536:str:0",\r\n10: "0:pint:0",\r\n11: "0:pint:0"\r\n'  # no idea
#    elif command is 'GetCommonSetting':
#      print 'I dunno'
#      payload += 'do me'
#    elif command is 'GetProjectCustomizationData':
#      print 'This function dumps the database schema.  Going to suck to parse the data.'
#      payload += 'do me'
    elif command is 'GetBugValue':
      # data={'id': '1,2,3,4', 'fields': ['BG_SEVERITY', 'BG_USER_02'], 'search': {'field1': 'value', 'field2': 'value2', 'limit': 1000}}	
      data['id']
      data['fields']
      #if not isinstance(data['id'], int):
      #  raise KeyError
      if isinstance(data['fields'], dict):
        data['fields'] = data['fields'].keys()
      elif not isinstance(data['fields'], list):
        raise KeyError
      fields = ''
      for x in data['fields']:
        fields += '%s,' % x
      fields = fields[0:len(fields)-1] # drop the last comma
      try: # if its not there, don't worry.  just create it as empty for the reference later
        data['search']
      except:
        data['search'] = {}
      payload += '0: \\%s\\%s,\r\n' % (self.qclen('0:conststr:%s' % command), '0:conststr:%s' % command)
      payload += '1: "0:int:%s",\r\n' % self.commandcount
      payload += '2: "0:int:%s",\r\n' % self.loginsessionid
      payload += '3: \\%s\\%s,\r\n' % (self.qclen('0:conststr:' + self.sessionkey), '0:conststr:' + self.sessionkey)
      payload += '4: "0:int:%s",\r\n' % self.projectsessionid

      query = '0:conststr:{\r\n'
      query2 = '[Filter]{\r\n'
      query2 += 'TableName:BUG,\r\n'
      query2 += 'ColumnName:BG_BUG_ID,\r\n'
      if data['search']:
        query2 += 'SortOrder:1,\r\n'
        query2 += 'SortDirection:1,\r\n'
      if data['id']:
        query2 += 'LogicalFilter:%s,\r\n' % data['id']
      query2 += 'NO_CASE:\r\n}\r\n'
# FIXME test this
      for key in data['search'].keys():
        query2 += '{\r\nTableName:BUG,\r\n'
        query2 += 'ColumnName:%s,\r\n' % key
        query2 += 'LogicalFilter:\\%s\\%s,\r\n' % (self.qclen(data['search'][key]), data['search'][key])
        query2 += 'NO_CASE:\r\n}\r\n'
        
      query += 'IDS:\\%s\\%s,\r\n' % (self.qclen(query2), query2)
      query += 'FIELDS:\\%s\\%s,\r\n' % (self.qclen(fields), fields)
      if data['limit']:
        query += 'LIMIT:%s,\r\n' % data['limit']
      query += 'HAS_LINKAGE:,\r\n'
      query += 'INCLUDE_RELATED_DATA:\r\n'
      query += '}\r\n'
      payload += '5: \\%s\\%s,\r\n' % (self.qclen(query), query)
      payload += '6: "65536:str:0"\r\n'
    elif command is 'PostBug':
      # data={'id': 12345, 'fields':{'key': 'newvalue', 'key2': 'newvalue2'}, 'oldvalues':{'key': 'oldvalue', 'key2': 'oldvalue2'}}
      # just going to pull from self.id and self.data for now
      try:
        data['id']
        data['fields']
        fields = []
        if not isinstance(data['id'], int):
          raise KeyError
        if not isinstance(data['fields'], dict):
          raise KeyError
      except KeyError as e:
        logging.exception('Missing or incorrect input for QC._buildpayload(\'%s\').  See QC._buildpayload.__doc__ : %s' % (command,e))
        raise
      except:
        logging.exception('General fail in QC._buildpayload(%s).' % command)
        raise
      payload += '0: \\%s\\%s,\r\n' % (self.qclen('0:conststr:%s' % command), '0:conststr:%s' % command)
      payload += '1: "0:int:%s",\r\n' % self.commandcount
      payload += '2: "0:int:%s",\r\n' % self.loginsessionid
      payload += '3: \\' + self.qclen('0:conststr:' + self.sessionkey) + '\\0:conststr:' + self.sessionkey + ',\r\n'
      payload += '4: "0:int:%s",\r\n' % self.projectsessionid
      query = '0:conststr:{\r\n'
      #if data['id'] is 0:
      if self.id is 0:
        query += 'METHOD:CREATE,\r\n' #new bug
# set default values for new bugs here.  Then merge in the defaults with overrides in data[].
        try:
          curdate = time.localtime()
          curdate = '%s-%s-%s' % (curdate[0],curdate[1],curdate[2])
          for key in self.default_new_case_fields.keys():
            try:
              #FIXME we should not use self.data directly in this function
              self.data[key]
            except:
              self.data[key] = self.default_new_case_fields[key]
        except:
          logging.exception('Failed to generate and merge New Bug data...')
          raise
      else:
        query += 'METHOD:POST,\r\n'
        query += 'ID:%s,\r\n' % self.id
      query += 'NEW_VALUES:\\'
      
      query2 = '{\r\n'
      for key in data['fields'].keys():
        if len(self.data[key]) > 1:
          query2 += '%s:\\%s\\%s,\r\n' % (key, self.qclen(self.data[key]), self.data[key])
        else:
          query2 += '%s:%s,\r\n' % (key, self.data[key])
      query2 = query2[0:len(query2)-3] # drop last ',\r\n'
      query2 += '\r\n}'
      query += '%s\\%s\r\n,\r\n' % (self.qclen(query2), query2)
      if data['oldvalues']:
        query += 'OLD_VALUES:\\'
        oldvalues = '{\r\n'
        for key in data['oldvalues'].keys():
          if len(data['oldvalues'][key]) > 1:
            oldvalues += '%s:\\%s\\%s,\r\n' % (key, self.qclen(data['oldvalues'][key]), data['oldvalues'][key])
          else:
            oldvalues += '%s:%s,\r\n' % (key, data['oldvalues'][key])
        oldvalues = oldvalues[0:len(oldvalues)-3] # drop last ',\r\n'
        oldvalues += '\r\n}'
        #oldvalues = 'OLD_VALUES:\\(qclen)\\{\r\nfield,oldvalue\r\n}\r\n'   # Too big of a pain to do this for real
        query += '%s\\%s\r\n,\r\n' % (self.qclen(oldvalues), oldvalues)
      query += 'UNLOCK:1\r\n}\r\n'
      payload += '5: \\%s\\%s,\r\n' % (self.qclen(query), query)
      payload += '6: "65536:str:0"\r\n'
    elif command is 'MailEntity':
      # data={'mailto': 'bugowner', 'cclist': 'cclist,addresses,list', 'id': 12345, 'fields':{'key': 'newvalue', 'key2': 'newvalue2'}}
# FIXME needs some input validation
      payload += '0: \\%s\\%s,\r\n' % (self.qclen('0:conststr:%s' % command), '0:conststr:%s' % command)
      payload += '1: "0:int:%s",\r\n' % self.commandcount
      payload += '2: "0:int:%s",\r\n' % self.loginsessionid
      payload += '3: \\%s\\%s,\r\n' % (self.qclen('0:conststr:' + self.sessionkey), '0:conststr:' + self.sessionkey)
      payload += '4: "0:int:%s",\r\n' % self.projectsessionid
      mailmsg = '{\r\n'
      mailmsg += 'TYPE:BUG,\r\n'
      mailmsg += 'SendTo:\\%s\\%s,\r\n' % (self.qclen(data['mailto']), data['mailto'])
      mailmsg += 'SendCc:'
      if data['cclist']:
        mailmsg += '\\%s\\%s' % (self.qclen(data['cclist']), data['cclist'])
      mailmsg += ',\r\n'
      mailmsg += 'FILTER:%s,\r\n' % data['id']
      subject = '[QC %s\\%s] Notification: Mailing To CC list: ID %s' % (self.qc_domainname,self.qc_projectname,data['id'])
      mailmsg += 'Subject:\\%s\\%s"",\r\n' % (self.qclen(subject), subject) 
      comment = 'This Item has been updated, see history to view changes'
      mailmsg += 'Comment:\\%s\\%s,\r\n' % (self.qclen(comment), comment)
      mailmsg += 'History:Y,\r\n'
      mailmsg += 'SingleMail:Y\r\n'
      mailmsg += '}\r\n'
      payload += '5: \\%s\\%s,\r\n' % (self.qclen(mailmsg), mailmsg)
      payload += '6: "65536:str:0"\r\n'
    elif command is 'DisconnectProject':
      payload += '0: \\%s\\%s,\r\n' % (self.qclen('0:conststr:%s' % command), '0:conststr:%s' % command)
      payload += '1: "0:int:%s",\r\n' % self.commandcount
      payload += '2: "0:int:%s",\r\n' % self.loginsessionid
      payload += '3: \\%s\\%s,\r\n' % (self.qclen('0:conststr:' + self.sessionkey), '0:conststr:' + self.sessionkey)
      payload += '4: "0:int:%s"\r\n' % self.projectsessionid
    elif command is 'Logout':
      payload += '0: \\%s\\%s,\r\n' % (self.qclen('0:conststr:%s' % command), '0:conststr:%s' % command)
      payload += '1: "0:int:%s",\r\n' % self.commandcount
      payload += '2: "0:int:%s",\r\n' % self.loginsessionid
      payload += '3: \\%s\\%s,\r\n' % (self.qclen('0:conststr:' + self.sessionkey), '0:conststr:' + self.sessionkey)
      payload += '4: "0:int:%s"\r\n' % self.projectsessionid
    elif command is 'ObjectLock':
      #data{'id': 12345, 'fields': data_dict.keys()}
      try:
        data['id']
        data['fields']
        if not isinstance(data['id'], int):
          raise KeyError
        if isinstance(data['fields'], list):
          fields = ''
          for x in data['fields']:
            fields += '%s,' % x
          fields = fields[0:len(fields)-1]  #chop the last comma off
        else:
          raise KeyError
      except KeyError as e:
        logging.exception('Missing or incorrect input for QC._buildpayload(\'%s\').  See QC._buildpayload.__doc__ : %s' % (command,e))
        raise
      except:
        error('General fail in QC._buildpayload(%s).' % command)
      payload += '0: \\%s\\%s,\r\n' % (self.qclen('0:conststr:%s' % command), '0:conststr:%s' % command)
      payload += '1: "0:int:%s",\r\n' % self.commandcount
      payload += '2: "0:int:%s",\r\n' % self.loginsessionid
      payload += '3: \\%s\\%s,\r\n' % (self.qclen('0:conststr:' + self.sessionkey), '0:conststr:' + self.sessionkey)
      payload += '4: "0:int:%s",\r\n' % self.projectsessionid
      lock = '0:conststr:\{\r\n'
      lock += 'TYPE:BUG,\r\n'
      lock += 'KEY:%s,\r\n' % data['id']
      lock += 'version:1,\r\n'
      lock += 'FIELDS:\\%s\\%s\r\n' % (self.qclen(fields), fields)
      lock += '}\r\n'
      payload += '5: \\%s\\%s,\r\n' % (self.qclen(lock),lock)
      payload += '6: "65536:str:0"\r\n'
    elif command is 'ObjectUnlock':
      #data{'id': '12345'}
      try:
        data['id']
        if not isinstance(data['id'], int):
          raise KeyError
      except KeyError as e:
        logging.exception('Missing or incorrect input for QC._buildpayload(\'%s\').  See QC._buildpayload.__doc__ : %s' % (command,e))
        raise
      except:
        logging.exception('General fail in QC._buildpayload(%s).' % command)
        raise
      payload += '0: \\%s\\%s,\r\n' % (self.qclen('0:conststr:%s' % command), '0:conststr:%s' % command)
      payload += '1: "0:int:%s",\r\n' % self.commandcount
      payload += '2: "0:int:%s",\r\n' % self.loginsessionid
      payload += '3: \\%s\\%s,\r\n' % (self.qclen('0:conststr:' + self.sessionkey), '0:conststr:' + self.sessionkey)
      payload += '4: "0:int:%s",\r\n' % self.projectsessionid
      lock = '0:conststr:\{\r\n'
      lock += 'TYPE:BUG,\r\n'
      lock += 'KEY:%s\r\n' % data['id']
      lock += '}\r\n'
      payload += '5: \\%s\\%s,\r\n' % (self.qclen(lock),lock)
      payload += '6: "65536:str:0"\r\n'
    elif command is 'ConnectProject':
      #data={'domainname': 'zbn_main', 'projectname': 'zbn'}
      try:
        data['domainname']
        data['projectname']
        if not isinstance(data['domainname'], str):
          raise KeyError
        if not isinstance(data['projectname'], str):
          raise KeyError
      except KeyError as e:
        logging.exception('Missing or incorrect input for QC._buildpayload(\'%s\').  See QC._buildpayload.__doc__ : %s' % (command,e))
        raise
      except:
        logging.exception('General fail in QC._buildpayload(%s).' % command)
        raise
      payload += '0: \\%s\\%s,\r\n' % (self.qclen('0:conststr:%s' % command), '0:conststr:%s' % command)
      payload += '1: "0:int:%s",\r\n' % self.commandcount
      payload += '2: "0:int:%s",\r\n' % self.loginsessionid
      payload += '3: "0:conststr:",\r\n' # no self.sessionkey yet
      payload += '4: "0:int:%s",\r\n' % self.projectsessionid
      projects = '0:conststr:{\r\n'
      projects += 'DOMAIN_NAME:%s,\r\n' % data['domainname']
      projects += 'PROJECT_NAME:%s\r\n' % data['projectname']
      projects += '}\r\n'
      payload += '5: \\%s\\%s,\r\n' % (self.qclen(projects), projects)
      payload += '6: "65536:str:0",\r\n'
      payload += '7: "0:pint:0"\r\n'  # why is this suddenly here?
    elif command is 'GetServerSettings':
      payload += '0: \\%s\\%s,\r\n' % (self.qclen('0:conststr:%s' % command), '0:conststr:%s' % command)
      payload += '1: "0:int:%s",\r\n' % self.commandcount
      payload += '2: "0:int:%s",\r\n' % self.loginsessionid
      payload += '3: "0:conststr:",\r\n' # no self.sessionkey yet
      payload += '4: "0:int:%s",\r\n' % self.projectsessionid
      payload += '5: "65536:str:0"\r\n'
    else:
      logging.error('Fail, unknown case-sensitive command.')
    payload += '}\r\n'

    #debug('Payload is: %s' % payload,name=self.name)
    return payload


  def password_encrypt(self,password):
    '''Encrypts your password under their fucking stupid algorithm.  This is meant to be called from QC.https_login(), unless, you know, you want to directly.

QC.password_encrypt('yourpassword')
'''

    result = ''
    resulta = []
    key = 'SmolkaWasHereMonSher'

    for i in range(0, len(password)):
      resulta.append(ord(password[i]) + ord(key[i % 20]))

    for i in range(0, len(resulta)):
      result += '%s!' % resulta[i]

    return 'ENRCRYPTED%s' % result


  def _get_td_id(self,param):
    '''This is the most retarded way to do data hashing, but that's how QC does it.  We had to reverse the Delphi DLL's to find the guid.'''
    
    ''' This only works for QC9
    td_id = 2301
    for i in range(0,len(input)):
      td_id += ord(input[i])

    return str(td_id)
'''
    guid = '{4947B489-F1D3-40e2-BD95-42851DC75CE6}'
    try:
      # py3k+:
      return hashlib.sha256(bytes(guid + param,encoding="UTF-8")).hexdigest().upper()
    except TypeError:
      #pre-py3k:
      return hashlib.sha256(guid + param).hexdigest().upper()


  def http_connect(self,username='',password=''):
    '''
# login url: baseurl + "/servlet/tdservlet/TDAPI_GeneralWebTreatment"
  call http_call to login
  
'''
    '''QC().http_connect(username=%s,password=%s)' % (username,password)'''
    #if not self.isrunning: error('Thread stopping.')
    try:
      if self.sessionkey:
        logging.debug("You're already connected.")
        return
    except:
      pass

    if not username: username = self.username
    if not password: password = self.password

    logging.info('Connecting to %s.' % self.qc_server)

    result = self._http_call(params=self._buildpayload('Login', {'username': username, 'password': password}))

    try:
      self.loginsessionid = re.findall('(?s)\r\nLOGIN_SESSION_ID:([^\x2c]+),', result[0])[0]
      self.sessionkey = re.findall('(?s)\r\nLOGIN_SESSION_KEY:\\x5c[0-9A-Fa-f]{8}\\x5c([^\x2c]+),', result[0])[0]
    except IndexError:
      logging.exception("Bad Username or Password.")
      raise

    result = self._http_call(params=self._buildpayload('ConnectProject', data={'domainname': self.qc_domainname, 'projectname': self.qc_projectname}))

    self.projectsessionid = re.findall('(?s)\r\nPROJECT_SESSION_ID:([^\x2c]+),', result[0])[0]

    logging.info('Connected to %s.' % self.qc_server)

    return


  def http_disconnect(self):
    '''QC.http_disconnect() - disconnects cleanly from QC'''
    if not self.sessionkey:
      return
    logging.debug('Disconnecting.')
    self._http_call(params=self._buildpayload(command='DisconnectProject'))
    self.projectsessionid = '-1'
    self._http_call(params=self._buildpayload(command='Logout'))
    self.sessionkey = None
    return


  def _buildheaders(self,headers={}):
    ''' builds headers for _http_call '''
    try:
      headers['User-Agent']
    except:
      headers['User-Agent'] = self.user_agent
    try:
      headers['Accept-Encoding']
    except:
      headers['Accept-Encoding'] = 'gzip'
    try:
      headers['Host']
    except:
      headers['Host'] = self.qc_server
    try:
      headers['Content-Type']
    except:
      headers['Content-Type'] = 'text/html; charset=UTF-8'
    try:
      headers['Connection']
    except:
      headers['Connection'] = 'Keep-Alive'
    try:
      headers['Cookie']
    except:
      if self.cookie:
        headers['Cookie'] = self.cookie
    return headers


  def _http_call(self, method='POST', url='', headers={}, params=''):
    ''' QC()._http_call(<input>) '''
    #if not self.isrunning: error('Thread stopping.')
    try:
      self.sessionkey
    except:
      logging.error("You need to log on first.")
      raise

# login url: baseurl + "/servlet/tdservlet/TDAPI_GeneralWebTreatment"

    if not url:
      url = '%s%s' % (self.qc_baseurl,'/servlet/tdservlet/TDAPI_GeneralWebTreatment')

    headers['X-TD-ID'] = self._get_td_id(params)

    headers = self._buildheaders(headers)

    logging.debug('request to be made: %s %s %s %s' % (method, url, params, headers))
    self.breakpoint(self._bp)

    conn = None

    try:
      if self.use_ssl:
        if self.cert != '':
          conn = httplib.HTTPSConnection(self.qc_server, cert_file=self.cert)
        else:
          conn = httplib.HTTPSConnection(self.qc_server)
      else:
        conn = httplib.HTTPConnection(self.qc_server)

      conn.request(method, url, params, headers)
      resp1 = conn.getresponse()

    except:
      logging.error('Unable to connect to %s (network failure).' % self.qc_server)
      raise

    logging.debug('HTTP response code: %s' % resp1.status)

    if resp1.getheader('Set-Cookie'):
      self.cookie = resp1.getheader('Set-Cookie')
      logging.debug('HTTP cookie: %s' % self.cookie)

    if resp1.status != 200:
      logging.error("Server did not respond with 200 OK, said: %s" % resp1.status)
      raise

    try:
      resp1.read(0) #won't return anything anyways
    except:
      logging.exception('Unable to connect to %s (data read error).' % self.qc_server)
      raise

    response = ''
    responses = []
    resp2 = ''
    
    try:
      import StringIO
      response = gzip.GzipFile(fileobj=StringIO.StringIO(resp1.read())).read()
      self._debug = response
    except (ImportError, NameError):
      # must be py3k
      response = gzip.decompress(resp1.read()).decode()
      self._debug = response
    except ValueError as e: 
      logging.error('Unable to connect to %s (data response error): %s.' % (self.qc_server,e))

    # most-verbose
    #logging.debug('Raw response is: %s' % response)

    #massage data
    if response:
      response = re.sub('(?s)^[^\\x7b]+\\x7b(.*)\\x7d\\x2c?\\x0d?\\x0a?$', '\\1', response)
      responses = re.split('(?s)(?:^|,)\r\n[0-9]:', response)[1:] # [1:] to skip empty [0] case
    else:
      logging.debug('Unable to read response. (empty response)')
      
    self.commandcount += 1

    self._debug = responses

    logging.debug('Responses is: %s' % responses)
    self.breakpoint(self._bp)

    return responses


if __name__ == '__main__':
  print('''This script is not ment to be ran by itself, but rather as a module.

Use pydoc to read documentation:

import pydoc
help('qcapi')
''')
  sys.exit(0)

import httplib2
import os
import argparse

from apiclient import discovery
import oauth2client
from oauth2client import file, client, tools

from ss import SSArea



class GoogleService:
  def __init__(self, appName, flags):
    # If modifying these scopes, delete your previously saved credentials
    # at ~/.credentials/sheets.googleapis.com-python-quickstart.json
    #SCOPES = 'https://www.googleapis.com/auth/spreadsheets.readonly'
    self.SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
    self.CLIENT_SECRET_FILE = 'client_secret.json'
    self.CREDENTIAL_FILE = 'credentials.json'
    self.APPLICATION_NAME = appName
    self.flags = flags
    #try:
    #  self.flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
    #except ImportError:
    #  self.flags = None

  def get_credentials(self):
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir, self.CREDENTIAL_FILE)

    store = file.Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(self.CLIENT_SECRET_FILE, self.SCOPES)
        flow.user_agent = self.APPLICATION_NAME
        if self.flags:
            credentials = tools.run_flow(flow, store, self.flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run_flow(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

  def get_service(self):
    credentials = self.get_credentials()
    http = credentials.authorize(httplib2.Http())

    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets', 'v4', http=http,
                              discoveryServiceUrl=discoveryUrl)
    return service



#class CSheetTab:
#  def __init__(self, service):
#    pass

class GRange:
  def __init__(self, tabName, rowstart=None, colstart=None, rowend = None, colend = None):
    self.tabName = tabName
    self.rowstart = rowstart
    self.colstart = colstart
    self.rowend = rowend
    self.colend = colend

  def nullEndCell(self):
    self.rowend = None
    self.colend = None

  def setStartCellName(self, name):
    (rowidx, colidx) = GRange.rangeIdx(name)
    self.setStartCellIdx(rowidx, colidx)

  def setEndCellName(self, name):
    (rowidx, colidx) = GRange.rangeIdx(name)
    self.setEndCellIdx(rowidx, colidx)

  # assume row/col-start are set
# set row/col-end to start + area size
  def setRangeArea(self, sarea):
    self.rowend = self.rowstart + sarea.rows - 1
    self.colend = self.colstart + sarea.cols - 1

  def setStartCellIdx(self, rowidx, colidx):
    self.rowstart = rowidx
    self.colstart = colidx

  def setEndCellIdx(self, rowidx, colidx):
    self.rowend = rowidx
    self.colend = colidx

  def name(self):
    r = "'" + self.tabName + "'!"
    r += GRange.cellName(self.rowstart, self.colstart)
    if self.rowend != None and self.colend != None:
      r += ':' + GRange.cellName(self.rowend , self.colend)
    return r

  # in:   0 based idx
  # out:  A-ZZ col name 
  @staticmethod
  def colName(colidx):
    _colName = ''
    while True:
      _colName = chr(colidx%26 + ord('A')) + _colName
      if colidx >= 26:
        colidx = colidx/26
      else:
        break
    return _colName
  
  # in:   0 based idx
  # out:  range, like A1
  @staticmethod
  def cellName(rowidx, colidx):
    _colName = GRange.colName(colidx)
    return _colName + str(rowidx + 1)

  # in:   range, like AB11
  # out:  (rowidx, colidx)
  @staticmethod
  def rangeIdx(range):
    colPart=True
    rowidx=0
    colidx=0
    for c in range:
      if colPart:
        if ord(c) >= ord('A') and ord(c) <= ord('Z'):
          colidx *= 26
          colidx += ord(c) - ord('A') + 1
        else:
          colPart = False
      if not colPart:
        if ord(c) < ord('0') or ord(c) > ord('9'):
          break
        rowidx *= 10
        rowidx += ord(c) - ord('0')
    colidx -= 1
    rowidx -= 1
    return (rowidx, colidx)




class GSheet:
  def __init__(self, service, spreadsheetId):
    self.service = service
    self.spreadsheetId = spreadsheetId
    self.clearBatchUpdateRequests()
    self.clearValuesUpdateRequests()
    self.properties = None


  def getValues(self, _range):
    result = self.service.spreadsheets().values().get(
        spreadsheetId=self.spreadsheetId, range=_range).execute()
    values = result.get('values', [])
    ss = None
    if values:
      cols = 0
      for row in values:
        if len(row) > cols:
          cols = len(row)
        #print("A cols = %d, row=%s " % (cols, str(row)))
      
      sarea = SSArea(cols)
      for row in values:
        x = int(cols - len(row))
        #print("B cols = %d, len(row)=%d, x=%d, row=%s " % (cols, len(row), x, str(row)))
        if len(row) < cols:
          for i in range(x):
            row.append('')
        sarea.addRow(row)

    return sarea

  def clearBatchUpdateRequests(self):
    self.batchUpdateRequests = \
    {
        "requests": []
    }

  def clearValuesUpdateRequests(self):
    self.valuesUpdateRequests = \
    {
      "valueInputOption": "USER_ENTERED",
      "data": [ ]
    }

  def getProperties(self):
    self.properties = self.service.spreadsheets().get(spreadsheetId=self.spreadsheetId).execute()

  def getSheetId(self, _sheetName):
    if self.properties is None:
      return None
    for sheet in self.properties['sheets']:
      if sheet['properties']['title'] == _sheetName:
        return sheet['properties']['sheetId']
    return None

  def getSheetIndex(self, _sheetName):
    if self.properties is None:
      return None
    for sheet in self.properties['sheets']:
      if sheet['properties']['title'] == _sheetName:
        return sheet['properties']['index']
    return None

  def getSheetByPos(self, _pos):
    if self.properties is None or _pos>=len(self.properties['sheets']):
      return None
    return self.properties['sheets'][_pos]
    
  def getSheetCount(self):
    if self.properties is None:
      return None
    return len(self.properties['sheets'])

  def execute(self):
    response=[None, None]
    if len(self.batchUpdateRequests["requests"]) > 0:
      response[0] = self.service.spreadsheets().batchUpdate(spreadsheetId=self.spreadsheetId,
                                               body=self.batchUpdateRequests).execute()
      self.clearBatchUpdateRequests()

    if len(self.valuesUpdateRequests["data"]) > 0:
      response[1] = self.service.spreadsheets().values().batchUpdate(spreadsheetId=self.spreadsheetId,
                                               body=self.valuesUpdateRequests).execute()
      self.clearValuesUpdateRequests()
    return response



  def insertPage(self, _title, _tabIdx = None):
    """
     AddSheetRequest
    """
    body = \
    {
      "addSheet":
      {
        "properties": 
        {
          "title": _title,
          #"index": number,
          "sheetType": "GRID",
          "gridProperties": {
            #"rowCount": 20,
            #"columnCount": 12
          },
          "hidden": "false",
          "tabColor": {
            "red": 1.0,
            "green": 0.3,
            "blue": 0.4
          },
          "rightToLeft": "false"
        }
      }
    }
    if _tabIdx != None:
      body["addSheet"]["properties"]["index"] = _tabIdx

    self.batchUpdateRequests["requests"].append(body)


  def resizePage(self, sheetId):
    body = \
    {
      "autoResizeDimensions": {
        "dimensions": {
          "sheetId": sheetId,
          "dimension": "COLUMNS",
          "startIndex": 0
          #opt  "endIndex": number
        }
      }
    }

    self.batchUpdateRequests["requests"].append(body)

  def addConditionalFormatRule(self, _sheetId, _startRow, _endRow, _startCol, _endCol, _formula, _colR, _colG, _colB, _ruleIdx):
    body = \
    {
      "addConditionalFormatRule": {
        "rule": {
          "ranges": [
            {
              "sheetId": _sheetId,
              "startColumnIndex": _startCol,
              "endColumnIndex": _endCol,
              "startRowIndex": _startRow,
              "endRowIndex": _endRow
            }
          ],
          "booleanRule": {
            "condition": {
              "type": "CUSTOM_FORMULA",
              "values": [
                {
                  "userEnteredValue": _formula
                }
              ]
            },
            "format": {
              "backgroundColor" : {
                "red": _colR,
                "green": _colG,
                "blue": _colB
              }
            }
          }
        }
        , "index": _ruleIdx
      }
    }

    self.batchUpdateRequests["requests"].append(body)


  def addHeaderStyle1(self, _sheetId, _startRow, _endRow):
    _colR = 0.71764
    _colG = 0.88235
    _colB = 0.80392

    body = \
    {
      "repeatCell": {
        "range": {
          "sheetId": _sheetId,
          "startRowIndex": _startRow,
          "endRowIndex": _endRow
        },
        "cell": {
          "userEnteredFormat": {
            "backgroundColor": {
              "red": _colR,
              "green": _colG,
              "blue": _colB
            },
            "horizontalAlignment" : "CENTER",
            "textFormat": {
              #"foregroundColor": {
              #  "red": 1.0,
              #  "green": 1.0,
              #  "blue": 1.0
              #},
              #"fontSize": 12,
              "bold": True
            }
          }
        },
        "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"
      }
    }

    self.batchUpdateRequests["requests"].append(body)

  def addHeaderStyle2(self, _sheetId, _startRow, _endRow):
    body = \
    {
      "repeatCell": {
        "range": {
          "sheetId": _sheetId,
          "startRowIndex": _startRow,
          "endRowIndex": _endRow
        },
        "cell": {
          "userEnteredFormat": {
            #"backgroundColor": {
            #  "red": _colR,
            #  "green": _colG,
            #  "blue": _colB
            #},
            #"horizontalAlignment" : "CENTER",
            "textFormat": {
              #"foregroundColor": {
              #  "red": 1.0,
              #  "green": 1.0,
              #  "blue": 1.0
              #},
              #"fontSize": 12,
              "bold": True
            }
          }
        },
        #"fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"
        "fields": "userEnteredFormat(textFormat)"
      }
    }

    self.batchUpdateRequests["requests"].append(body)


  def freezeRows(self, _sheetId, _rows):
    body = \
    {
      "updateSheetProperties": {
        "properties": {
          "sheetId": _sheetId,
          "gridProperties": {
            "frozenRowCount": _rows
          }
        },
        "fields": "gridProperties.frozenRowCount"
      }
    }

    self.batchUpdateRequests["requests"].append(body)

  def freezeCols(self, _sheetId, _cols):
    body = \
    {
      "updateSheetProperties": {
        "properties": {
          "sheetId": _sheetId,
          "gridProperties": {
            "frozenColumnCount": _cols
          }
        },
        "fields": "gridProperties.frozenColumnCount"
      }
    }

    self.batchUpdateRequests["requests"].append(body)

  def setError(self, _range, _type, _text):
    """
     type = ERROR_TYPE_UNSPECIFIED, ERROR, LOADING, ...
    UpdateCellsRequest

    """

    rowidx = 1
    colidx = 10

    body = \
    {
      "updateCells":
      {
        "rows": [
          {
            "values": [
              {
                "effectiveValue": 
                {
                  "errorValue": {
                    "type": _type,
                    "message": _text
                  }
                }
              }
            ]
          }
        ],
        "start": {
          "sheetId": 0,
          "rowIndex": rowidx,
          "columnIndex": colidx
        },
        "fields": "*"
      }
    }
    self.batchUpdateRequests["requests"].append(body)

  def deleteSheet(self, _sheetId):
    body = \
    {
      "deleteSheet": 
      {
        "sheetId": _sheetId
      }
    }
    self.batchUpdateRequests["requests"].append(body)
    


  def setValues(self, _range, _sarea):
    data = \
    {
      "range": _range,
      "majorDimension": "ROWS",
      "values": []
    }
    data["values"] = _sarea.toVector()

    self.valuesUpdateRequests["data"].append(data)


  def setCell(self, _range, _value):
    data = \
    {
      "range": _range,
      "majorDimension": "ROWS",
      "values": []
    }
    data["values"] = [[_value]]

    self.valuesUpdateRequests["data"].append(data)


  def __setitem__(self, i, v):
    self.data[i] = v

  def __len__(self):
    return self.cols



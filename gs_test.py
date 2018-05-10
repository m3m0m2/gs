#!/usr/bin/python

from gs import *
from ss import SSArea




def testRange():
  sarea = SSArea(3, 2)
  sarea[0][0] = 1
  sarea[1][1] = 2
  sarea[0][2] = 3
  sarea[1][2] = 4

  grange = GRange("Update 2016-10-01 13:14:15", 0, 0)
  grange.setEndCellName('D3')
  print("grange1: " + grange.name())

  grange.setStartCellIdx(4, 4)
  grange.setRangeArea(sarea)
  print("grange2: " + grange.name())
  print("sarea: " + str(sarea))

  grange.setStartCellIdx(0, 0)
  grange.setEndCellIdx(1019, 26)
  print("grange3: " + grange.name())

def testInsertTab():
  spreadsheetId = '1LvznVc5gqDDNFKo-Dhtc6OIDAahEI4GWzd2Qf1_E0jw'
  service = GoogleService().get_service()
  sheet = GSheet(service, spreadsheetId)
  #sheet.insertPage("Test Page", 1)
  sheet.insertPage("Test Page2")
  sheet.execute()

def testResizePage():
  spreadsheetId = '1LvznVc5gqDDNFKo-Dhtc6OIDAahEI4GWzd2Qf1_E0jw'
  service = GoogleService().get_service()
  sheet = GSheet(service, spreadsheetId)
  sheet.resizePage(0)
  sheet.execute()

def testSpreadsheet():
  spreadsheetId = '1LvznVc5gqDDNFKo-Dhtc6OIDAahEI4GWzd2Qf1_E0jw'
  service = GoogleService().get_service()
  sheet = GSheet(service, spreadsheetId)

  sarea = sheet.getValues('Sheet1!A6:D')

  print("### sarea")
  print(sarea) 

  sarea2 = sarea.frame(1, 0, 2, 3)
  print("### sarea2")
  print(sarea2) 

  sarea3 = sarea.frame(2, 1, 2, 3)
  print("### sarea3")
  print(sarea3) 

  sheet.setValues('Sheet1!H6', sarea2)
  sheet.setValues('Sheet1!M7', sarea3)
  response = sheet.execute()
  print("response: " + str(response))

def testGetProperties():
  spreadsheetId = '1LvznVc5gqDDNFKo-Dhtc6OIDAahEI4GWzd2Qf1_E0jw'
  service = GoogleService().get_service()
  sheet = GSheet(service, spreadsheetId)
  sheet.getProperties()
  print("response:\n" + str(sheet.properties))

def testConditionalRuleFormat():
  spreadsheetId = '1LvznVc5gqDDNFKo-Dhtc6OIDAahEI4GWzd2Qf1_E0jw'
  service = GoogleService().get_service()
  sheet = GSheet(service, spreadsheetId)
  sheet.getProperties()
  sheetId = sheet.getSheetId('Update_2016-10-25 15:42:27')
  """
  response = sheet.addConditionalFormatRule(_sheetId, _startRow, _endRow, _startCol, _endCol, _formula, _colR, _colG, _colB, _ruleIdx):
  """
  sheet.addConditionalFormatRule(sheetId, 3, 4, 4, 20, '=AND(NOT(ISBLANK(E4)),LTE(E4,12))', 0.95686275, 0.78039217, 0.7647059, 0)
  sheet.addConditionalFormatRule(sheetId, 3, 4, 4, 20, '=AND(NOT(ISBLANK(E4)),LT(E4,24),GT(E4,12))', 0.98823529, 0.90980393, 0.69803923, 1)
  response = sheet.execute()
  #print("response:\n" + str(response))

def testHeader():
  spreadsheetId = '1LvznVc5gqDDNFKo-Dhtc6OIDAahEI4GWzd2Qf1_E0jw'
  service = GoogleService().get_service()
  sheet = GSheet(service, spreadsheetId)
  sheet.getProperties()
  sheetId = sheet.getSheetId('Check_2016-10-27 18:01:51')
  sheet.addHeaderStyle1(sheetId, 2, 3)
  response = sheet.execute()
  print("response:\n" + str(response))

def main():
  #testSpreadsheet()
  #testRange()
  #testInsertTab()
  #testResizePage()
  #testGetProperties()
  #testConditionalRuleFormat()
  testHeader()


if __name__ == '__main__':
    main()

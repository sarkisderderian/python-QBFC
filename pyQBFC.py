#!usr/bin/python
import win32com.client
#No ElementTree needed, since no raw XML

# Open a QB Session
sessionManager = win32com.client.Dispatch("QBFC10.QBSessionManager")
sessionManager.OpenConnection('', 'Test QBFC Request')
# No ticket needed in QBFC
sessionManager.BeginSession("", 0)

# Send query and receive response
requestMsgSet = sessionManager.CreateMsgSetRequest("US", 6, 0)
requestMsgSet.AppendAccountQueryRq()
responseMsgSet = sessionManager.DoRequests(requestMsgSet)


#Peel away the layers of response
QBXML = responseMsgSet
QBXMLMsgsRq = QBXML.ResponseList
AppendAccountQueryRq = QBXMLMsgsRq.GetAt(0)
for x in range(0, len(AppendAccountQueryRq.Detail)):
    AccountRet= AppendAccountQueryRq.Detail.GetAt(x)
    name = AccountRet.Name.GetValue()
    balance = AccountRet.Balance.GetValue()
    print(name+"\t|\t"+str(balance))

# Disconnect from Quickbooks
sessionManager.EndSession()           # Close the company file (no ticket needed)
sessionManager.CloseConnection()      # Close the connection

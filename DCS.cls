VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'allows the DCS (Dll Cgi-Script) to communicate with the webserver
'includes many usefull commands. This is for dll cgi scripts, not perl
'perl cgi-scripts are run from form1

'buffer for return data
Private backdata As String

'the data after the ? (data 'GET')
Public getdata As String

'the data posted (data 'POST')
Public postdata As String

'set this to true when everything is done, and the results can be sent back
Public alldone As Boolean

Public Sub senddata(data)
    'the dcs calls this to send (some) data back to the host
    'note that the data is buffered, until the alldone call is sent
    'equiv of perl's 'print'
    backdata = backdata & data
End Sub

Public Sub reset()
    'destroys all sent data, usually used before an error is sent
    backdata = ""
End Sub

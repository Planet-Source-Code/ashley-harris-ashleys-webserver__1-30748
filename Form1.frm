VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ashley's Web Server"
   ClientHeight    =   855
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   855
   ScaleWidth      =   2310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Top             =   450
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   540
      Top             =   240
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   1545
      Top             =   135
   End
   Begin MSWinsockLib.Winsock ws 
      Index           =   0
      Left            =   1020
      Top             =   30
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   80
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   15
      TabIndex        =   2
      Top             =   585
      Width           =   3105
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   30
      TabIndex        =   1
      Top             =   315
      Width           =   3105
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   30
      TabIndex        =   0
      Top             =   15
      Width           =   3105
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const websitedir = "" 'would be something like C:\docroot\, blank means app.path
Private Const default = "test.html" 'can put '?a=b' etc. "" means directory listing
Private Const error404 = "/404.html"
Private docroot As String

Private fso As New FileSystemObject
Dim connections As Integer
Dim header As String
Dim sendbackindex As Integer

'only one cgi script can be executed at a time, all others are put in a que
Public executingcgi As Boolean

Private Sub Form_Load()
    'sets up everything
    ws(0).Listen
    StayOnTop Me, True
    docroot = IIf(websitedir = "", App.Path, websitedir)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'closes the socket
    ws(0).Close
End Sub

Private Sub Timer1_Timer()
    'displays how many connections have been made
    Label1.Caption = "connections: " & connections
End Sub

Private Sub Timer2_Timer()
    'return the results from the cgi script to the server
    'give it some time to write the output file.
    Dim filedata As String
        
    Open "C:\temp.txt" For Binary Access Read As #1
    filedata = Space$(LOF(1))
    Get #1, , filedata
    Close #1
        
    header = header & "Connection: close" & vbNewLine & "Transfer-Encoding: chunked" & vbNewLine
    ws(sendbackindex).senddata header & filedata
    Timer2 = False
    executingcgi = False
End Sub

Private Sub ws_ConnectionRequest(index As Integer, ByVal requestID As Long)
    'whenever someone requests a webpage, move their request to the side, so
    'we can accept more. (Just I've notice IE will request about 10 files at once)
    connections = connections + 1
    Load ws(connections)
    ws(connections).Accept requestID
End Sub

Private Sub ws_DataArrival(index As Integer, ByVal bytesTotal As Long)
    'The main sub, handles the the request, inc CGI and content-type asociations
    Dim a As String
    ws(index).getdata a
    If a = "" Then
        ws(index).Close
        Exit Sub
    End If
    
    'only the first line is relevant.
    a = Left(a, InStr(1, a, vbNewLine) - 1)
    a = Mid(a, InStr(1, a, " ") + 1)
    a = Left(a, InStr(1, a, " ") - 1)
    
    'display what file is being requested
    Label2.Caption = a
    
    'display to whom the file is being send
    Label3.Caption = ws(index).RemoteHostIP
    
    If Right(a, 1) = "/" Then a = a & default
    
    'seperated the request string into filename and GET data
    If Not CBool(InStr(1, a, "?")) Then
        a = a & "?"
    End If
    cmd = Left(a, InStr(1, a, "?") - 1)
    data = Mid(a, InStr(1, a, "?") + 1)
    cmd = Replace(cmd, "/", "\")
    cmd = Replace(cmd, "%20", " ")
    Path = fso.BuildPath(docroot, cmd)
    header = "HTTP/1.0 200 OK" & vbNewLine & "Server: A's webserver" & vbNewLine
    
    If default = "" And Right(cmd, 1) = "\" Then
        'do a directory listing
        header = header & "Connection: close" & vbNewLine & "Transfer-Encoding: chunked" & vbNewLine
        header = header & "Content-Type: " & "Text/html" & vbNewLine & vbNewLine
        
        ws(index).senddata header & dirlisting(CStr(Path), Replace(cmd, "\", "/"))
        Exit Sub
    End If
        
    If Not fso.FileExists(Path) Then
        'file doesn't exist, ok, so load 404 page
filenotfound:
        a = error404
        
        'reprocess all this
        If Not CBool(InStr(1, a, "?")) Then
            a = a & "?"
        End If
        
        cmd = Left(a, InStr(1, a, "?") - 1)
        data = Mid(a, InStr(1, a, "?") + 1)
        cmd = Replace(cmd, "/", "\")
        Path = fso.BuildPath(docroot, cmd)
        
        'should be HTTP/1.0 404 ERROR, but then IE didn't display it, so, just say 200 OK
        'header = "HTTP/1.0 200 OK" & vbNewLine & "Server: A's webserver" & vbNewLine
    End If
    
    
    'Is it a cgi script?
    Select Case Mid(Path, InStrRev(Path, ".") + 1)
    
    Case "cgi" 'standard perl scripts
        iscgi = True
    Case "pl"
        iscgi = True
        
    Case "dll" 'my own type of cgi script
        iscgi = True
        
    Case Else 'something non-cgi
        iscgi = False
    End Select
    
    If Not iscgi Then
        'just a plain file, ok, send it back.
        header = header & "Accept-Ranges: bytes" & vbNewLine
        On Error GoTo filenotfound
        header = header & "Content-Length: " & FileLen(Path) & vbNewLine
        On Error GoTo 0
        'Asign the extensions to the types.
        'course, if I've forgotten any, just add them
        Select Case LCase(Mid(Path, InStrRev(Path, ".") + 1))
        Case "html" 'HTML file
            cont = "text/html"
        Case "htm"  'HTML file
            cont = "text/html"
        Case "txt"  'TEXT (notepad) file
            cont = "text/text"
        Case "js"   'Javascript library
            cont = "text/html" 'YES, that is right
        Case "pdf"  'ADOBE ACROBAT PDF file
            cont = "application/pdf"
        Case "sit"  'STUFFIT archive
            cont = "application/x-stuffit"
        Case "avi"  'AUDIO VISUAL video
            cont = "video/avi"
        Case "css"  'CASSCADING STYLE SHEET formating info
            cont = "text/css"
        Case "swf"  'SHOCKWAVE FLASH animation
            cont = "application/futuresplash"
        Case "jpg"  'JOINT PHOTOGROPHERS EXPERT GROUP image
            cont = "image/jpeg"
        Case "xls"  'MICROSOFT EXCEL spreadsheet
            cont = "application/vnd.ms-excel"
        Case "doc"  'MICROSOFT WORD formated text
            cont = "application/vnd.ms-word"
        Case "midi" 'MUSICAL INSTRUMENT DIGITAL INTERFACE music
            cont = "audio/midi"
        Case "mp3"  'MOTION PICTURE EXPERT GROUP LAYER 3 music
            cont = "audio/mpeg"
        Case "rm"   'REAL MEDIA video
            cont = "application/vnd.rn-realmedia"
        Case "rtf"  'MICROSOFT RICHTEXT formatted text
            cont = "application/msword"
        Case "wav"  'WAVE sound
            cont = "audio/wav"
        Case "zip"  'ZIP archive
            cont = "application/x-tar"
        Case "png"  'PORTABLE NETWORK GRAPGHICS image
            cont = "image/png"
        Case "gif"  'COMPUSERVE GRAPHICS INTERGHANGE FORMAT Image
            cont = "image/gif"
        End Select
        Dim filedata As String
        
        'Load up the file
        Open Path For Binary Access Read As #1
        filedata = Space$(LOF(1))
        Get #1, , filedata
        Close #1

        'Finish the header
        header = header & "Connection: close" & vbNewLine
        header = header & "Content-Type: " & cont & vbNewLine & vbNewLine
        filedata = header & filedata
        
        'send the reply
        ws(index).senddata filedata
    ElseIf Left(Path, 4) = ".dll" Then
        'It's my own type of cgi-script! run it
        
    Else
        'Run perl CGI scripts, and pipe the result back to the server
        
        'create a temp perl script, calling the script in question, setting the
        'enviroment variables QUERRY_STRING and REMOTE_ADDR, and, anything else
        'I feel like doing
        
        While executingcgi
            DoEvents
        Wend
        executingcgi = True
        
        fn = docroot & "temp.pl"
        Open fn For Output As #2
        Print #2, "" & _
        "#! /usr/bin/perl" & vbNewLine & _
        "$ENV{QUERY_STRING} = '" & data & "';" & vbNewLine & _
        "$ENV{REMOTE_ADDR} = '" & ws(index).RemoteHostIP & "';" & vbNewLine & _
        "$ENV{SERVER_SOFTWARE} = '" & "A`s Webserver/1.1 (Windows)" & "';" & vbNewLine & _
        "$ENV{GATEWAY_INTERFACE} = '" & "CGI/GET ONLY" & "';" & vbNewLine & _
        "$ENV{DOCUMENT_ROOT} = '" & Replace(docroot, "\", "\\") & "';" & vbNewLine & _
        "$ENV{SERVER_PROTOCOL} = '" & "HTTP/1.0" & "';" & vbNewLine & _
        "$ENV{REQUEST_METHOD} = '" & "GET" & "';" & vbNewLine & _
        "$ENV{SERVER_ADDR} = '" & ws(0).LocalIP & "';" & vbNewLine & _
        "$ENV{SCRIPT_FILENAME} = '" & Replace(Path, "\", "\\") & "';" & vbNewLine & _
        "$ENV{SCRIPT_NAME} = '" & Replace(cmd, "\", "\\") & "';" & vbNewLine & _
        "$ENV{SERVER_NAME} = '" & ws(0).LocalHostName & "';" & vbNewLine & _
        "$ENV{SERVER_PORT} = '" & ws(0).LocalPort & "';" & vbNewLine & _
        "do '" & Path & "';"
        Close #2
        
        'go to the directory with the cgi script in it. (To avoid a bug in perl)
        ChDir Left(fn, InStrRev(fn, "\"))
        
        On Error Resume Next
        If fso.FileExists("C:\temp.txt") Then Kill "C:\temp.txt"
        On Error GoTo 0
        
        'run the perl script (PERL MUST BE IN YOUR AUTOEXEC PATH VARIABLE)
        Shell "command.com /c perl " & fso.GetFileName(fn) & " >""C:\temp.txt""", vbNormalFocus
        sendbackindex = index
        Timer2 = True
    End If
    
End Sub
    
Private Sub ws_Error(index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'something went wrong, about gracefully
    ws(index).Close
    If index = 0 Then ws(0).Listen Else Unload ws(index)
End Sub

Private Sub ws_SendComplete(index As Integer)
    'close when done
    ws(index).Close
    Unload ws(index)
End Sub


Public Function dirlisting(dir As String, url As String) As String
    If Right(url, 1) <> "/" Then url = url & "/"
    If Right(dir, 1) <> "\" Then dir = dir & "\"
    
    back = "<HTML><HEAD><TITLE>" & dir & "</TITLE></HEAD><BODY>" & vbNewLine
    If fso.FileExists(dir) Then
        back = back & "FOLDER: " & dir & " NOT FOUND"
        dirlisting = back
        Exit Function
    End If
    back = back & "<TABLE width=100%><tr><td><B>Name</B></td><td><B>Size (bytes)</B></td><td><B>Date</B></td></tr>"
    Dim x() As String
    x() = GetSubFolders(dir)
    On Error GoTo nosubdirs
    For a = 0 To UBound(x())
        fname = x(a)
        fsize = "-"
        fdate = "-"
        back = back & "<tr><td><A HREF=""" & url & fname & "\"">" & fname
        back = back & "</A></td><td>" & fsize & "</td><td>" & fdate
        back = back & "</td></tr>" & vbNewLine
    Next a
    On Error GoTo 0
nosubdirs:
    File1.Path = dir
    For a = 0 To File1.ListCount - 1
        fname = File1.List(a)
        fsize = FileLen(dir & fname)
        fdate = FileDateTime(dir & fname)
        back = back & "<tr><td><A HREF=""" & url & fname & """>" & fname
        back = back & "</A></td><td>" & fsize & "</td><td>" & fdate
        back = back & "</td></tr>" & vbNewLine
    Next a
    back = back & "</TABLE></BODY></HTML>"
    dirlisting = back
End Function

Function GetSubFolders(folder) As Variant
    Dim fnames() As String
    
   If Right(folder, 1) <> "\" Then folder = folder & "\"
   fd = dir(folder, vbDirectory)
   While fd <> ""
    If (GetAttr(folder & fd) And vbDirectory) = vbDirectory Then
        push fnames(), fd
    End If
    fd = dir()
   Wend
   GetSubFolders = fnames()
End Function

Private Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init: ReDim ary(0): ary(0) = value
End Sub

VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PutGet FTP"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   6465
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtFileR 
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Top             =   2760
      Width           =   4335
   End
   Begin VB.TextBox TxtFile 
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Top             =   2400
      Width           =   4335
   End
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   375
      Left            =   5040
      TabIndex        =   8
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Get"
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Top             =   960
      Width           =   4215
   End
   Begin VB.TextBox txtDir 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1920
      TabIndex        =   1
      Text            =   "/"
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Frame Status1 
      Caption         =   "Status"
      Height          =   735
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   5895
      Begin VB.Label Status 
         Caption         =   "Disconnected..."
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   5655
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Put"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Remote File Name:"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   6240
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label Label5 
      Caption         =   "Local File Name:"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Password:"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "User:"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Server:"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Initial Dirfectory:"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1320
      Width           =   1575
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private hOpen As Long
Private hConnection As Long
Private hFile As Long
Private dwType As Long
Private dwSeman As Long
Private Sub Command1_Click()
    
    'Variables
    sServer = txtServer.Text
    sUser = txtUser.Text
    sPassword = txtPassword.Text
    sDir = txtDir.Text
    sLocal = App.Path & "\" & TxtFile.Text
    sRemote = TxtFileR.Text
    
    'Save values to remember
    SaveSetting "PutGet FTP", "Values", "Server", txtServer.Text
    SaveSetting "PutGet FTP", "Values", "User", txtUser.Text
    SaveSetting "PutGet FTP", "Values", "Password", txtPassword.Text
    SaveSetting "PutGet FTP", "Values", "Directory", txtDir.Text
    
'Open INTERNET
    hOpen = InternetOpen("PutGet FTP", INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
    If hOpen = 0 Then 'ZERO means Internet Coudn't Open
        MsgBox "Error: " & Err.LastDllError, 32, "Internet Conection Error"
        Status.Caption = "Disconnected..."
        Exit Sub
    End If
    Status.Caption = "Internet Open..."
    
    dwType = FTP_TRANSFER_TYPE_BINARY 'SET TO BINARY
    dwSeman = 0 'Set Conection Active
    hConnection = 0 'Reset Conection
    
'Connect to server
    hConnection = InternetConnect(hOpen, sServer, INTERNET_INVALID_PORT_NUMBER, sUser, sPassword, INTERNET_SERVICE_FTP, dwSeman, 0)
    If hConnection = 0 Then 'ZERO means can't connect to Server
        MsgBox "Error: " & Err.LastDllError, 32, "Server Conection Error"
        Status.Caption = "Disconnected..."
        Exit Sub
    End If
    Status.Caption = "Connected to Server..."

'Specify Initial Directory
    OpenDir = FtpSetCurrentDirectory(hConnection, sDir)
    If OpenDir = False Then 'False means specified directory is wrong
        MsgBox "Error: " & Err.LastDllError, 32, "Initial Directory Error"
        Status.Caption = "Disconnected..."
        If hConnection <> 0 Then 'Disconnect if is still conected
            Cerrar = InternetCloseHandle(hConnection)
        End If
        Exit Sub
    End If
    Status.Caption = "Directory Ready..."

'Put File
    Subir = FTPPutFile(hConnection, sLocal, sRemote, dwType, 0)
    If Subir = False Then 'False means couldn't send the file
        MsgBox "Error: " & Err.LastDllError, 32, "File Transfer Error"
        Status.Caption = "Disconnected..."
        If hConnection <> 0 Then 'Disconnect if is still conected
            Cerrar = InternetCloseHandle(hConnection)
        End If
        Exit Sub
    End If
    Status.Caption = "Sending File..."
    
'Close conection
    If hConnection <> 0 Then
        Cerrar = InternetCloseHandle(hConnection)
        Status.Caption = "Disconnected..."
    End If
    
End Sub

Private Sub Command2_Click()
    'Variables
    sServer = txtServer.Text
    sUser = txtUser.Text
    sPassword = txtPassword.Text
    sDir = txtDir.Text
    sLocal = App.Path & "\Prueba.txt"
    sRemote = "Prueba.txt"
    
    'Save values to remember
    SaveSetting "PutGet FTP", "Values", "Server", txtServer.Text
    SaveSetting "PutGet FTP", "Values", "User", txtUser.Text
    SaveSetting "PutGet FTP", "Values", "Password", txtPassword.Text
    SaveSetting "PutGet FTP", "Values", "Directory", txtDir.Text
    
'Open INTERNET
    hOpen = InternetOpen("CYSM FTP", INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
    If hOpen = 0 Then 'ZERO means Internet Coudn't Open
        MsgBox "Error: " & Err.LastDllError, 32, "Internet Conection Error"
        Status.Caption = "Disconnected..."
        Exit Sub
    End If
    Status.Caption = "Internet Open..."
    
    dwType = FTP_TRANSFER_TYPE_BINARY 'SET TO BINARY
    dwSeman = 0 'Set Conection Active
    hConnection = 0 'Reset Conection

'Connect to server
    hConnection = InternetConnect(hOpen, sServer, INTERNET_INVALID_PORT_NUMBER, sUser, sPassword, INTERNET_SERVICE_FTP, dwSeman, 0)
    If hConnection = 0 Then 'ZERO means can't connect to Server
        MsgBox "Error: " & Err.LastDllError, 32, "Server Conection Error"
        Status.Caption = "Disconnected..."
        Exit Sub
    End If
    Status.Caption = "Connected to Server..."

'Specify Initial Directory
    OpenDir = FtpSetCurrentDirectory(hConnection, sDir)
    If OpenDir = False Then 'False means specified directory is wrong
        MsgBox "Error: " & Err.LastDllError, 32, "Error de Directorio Inicial"
        Status.Caption = "Disconnected..."
        If hConnection <> 0 Then 'Disconnect if is still conected
            Cerrar = InternetCloseHandle(hConnection)
        End If
        Exit Sub
    End If
    Status.Caption = "Directory Ready..."

'Get File
    Bajar = FTPGetFile(hConnection, sRemote, sLocal, False, FILE_ATTRIBUTE_NORMAL, dwType Or INTERNET_FLAG_RELOAD, 0)
    If Bajar = False Then 'False means couldn't get the file
        MsgBox "Error: " & Err.LastDllError, 32, "File Transfer Error"
        Status.Caption = "Disconnected..."
        If hConnection <> 0 Then 'Disconnect if is still conected
            Cerrar = InternetCloseHandle(hConnection)
        End If
        Exit Sub
    End If
    Status.Caption = "Downloading File..."
    
'Close conection
    If hConnection <> 0 Then
        Cerrar = InternetCloseHandle(hConnection)
        Status.Caption = "Disconnected..."
    End If
    

End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Form_Load()
'Values to remember

    txtServer.Text = GetSetting("PutGet FTP", "Values", "Server")
    txtUser.Text = GetSetting("PutGet FTP", "Values", "User")
    txtPassword.Text = GetSetting("PutGet FTP", "Values", "Password")
    txtDir.Text = GetSetting("PutGet FTP", "Values", "Directory")
    TxtFile.Text = "Test.txt"
    TxtFileR.Text = "Test.txt"
    
End Sub

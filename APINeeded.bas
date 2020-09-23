Attribute VB_Name = "APINeeded"
'---------------------- PutGet FTP ---------------------------
'THIS APP USES ONLY THE NECESARY API'S TO PUT OR GET A FILE
'FROM AN FTP SITE.

'IT'S USEFUL TO MAKE CONNECTIONS TO INTERNET SERVERS FROM CLIENT
'APPLICATIONS, FOR EXAMPLE, TO UPDATE INFORMATION ON A WEB SITE
'FROM AN APPLICATION OR TO BRING TO A LOCAL COMPUTER INFORMATION
'FROM THE INTERNET.

'I HOPE YOU FIND IT USEFUL...
'JMORENO - jmoreno@cysm.com.mx
'********************************************************

Public Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Public Const FTP_TRANSFER_TYPE_BINARY = &H2
Public Const FTP_TRANSFER_TYPE_ASCII = &H1
Public Const INTERNET_INVALID_PORT_NUMBER = 0
Public Const INTERNET_SERVICE_FTP = 1
Public Const INTERNET_FLAG_RELOAD = &H80000000
Public Const FILE_ATTRIBUTE_NORMAL = &H80

Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Public Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Public Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
Public Declare Function FTPPutFile Lib "wininet.dll" Alias "FtpPutFileA" (ByVal hFtpSession As Long, ByVal lpszLocalFile As String, ByVal lpszRemoteFile As String, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean
Public Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Long
Public Declare Function FTPGetFile Lib "wininet.dll" Alias "FtpGetFileA" (ByVal hFtpSession As Long, ByVal lpszRemoteFile As String, ByVal lpszNewFile As String, ByVal fFailIfExists As Boolean, ByVal dwFlagsAndAttributes As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean


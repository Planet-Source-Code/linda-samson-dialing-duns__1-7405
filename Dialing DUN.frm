VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dialing DUNs"
   ClientHeight    =   1740
   ClientLeft      =   2280
   ClientTop       =   2910
   ClientWidth     =   5400
   Icon            =   "Dialing DUN.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   5400
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   375
      Left            =   3750
      TabIndex        =   3
      Top             =   1290
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Connect"
      Default         =   -1  'True
      Height          =   375
      Left            =   3750
      TabIndex        =   2
      Top             =   900
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   180
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   480
      Width           =   5055
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "linda.69@mailcity.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   210
      TabIndex        =   4
      Top             =   1260
      Width           =   2340
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Choose an Internet Connection:"
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   120
      Width           =   2250
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'linda.69@mailcity.com
'List all current DUNs and option to connect to it.

Option Explicit

' Registry Functions...
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

' Registry constants...
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003

Const ERROR_SUCCESS = 0&

Const SYNCHRONIZE = &H100000
Const STANDARD_RIGHTS_READ = &H20000
Const STANDARD_RIGHTS_WRITE = &H20000
Const STANDARD_RIGHTS_EXECUTE = &H20000
Const STANDARD_RIGHTS_REQUIRED = &HF0000
Const STANDARD_RIGHTS_ALL = &H1F0000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))

Const REG_DWORD = 4
Const REG_BINARY = 3
Const REG_SZ = 1

Const MaxValue = 256
Const MaxLength = 20       'if this value is too short for you, increase it

Dim KeyHandle As Long
Dim KeyHandle2 As Long
Dim Result As Long
Dim CurrentIndex As Long
Dim tmpString As String
Dim userName As String
Dim connName As String

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub Command1_Click()
On Error Resume Next             'do we need this?

  'Extract Connection Name, remove all trailing blanks
  connName = RTrim(Mid(Combo1.Text, 1, MaxLength))
  Shell "rundll rnaui.dll,RnaDial " + connName, vbNormalFocus
End Sub

Private Sub Form_Load()
On Error Resume Next        'do we need this?
  
  Me.Top = 900
  Me.Left = 600
    
  ' Open...
  Result = RegOpenKeyEx(HKEY_CURRENT_USER, "RemoteAccess\Addresses", 0&, KEY_READ, KeyHandle)
    
  ' OK?
  If Result <> ERROR_SUCCESS Then
    MsgBox "An error has occurred opening the specified registry key.", vbCritical, "Error: Registry"
      Exit Sub
  End If

  ' we will start with first entry
  CurrentIndex = 0
   
  ' Start enumerating the DUN's...
  Do
    connName = Space(MaxValue)
    userName = Space(MaxValue)
    tmpString = Space(MaxValue)
    Result = RegEnumValue(KeyHandle, CurrentIndex, connName, MaxValue, 0&, REG_DWORD, ByVal 0&, MaxValue)
    CurrentIndex = CurrentIndex + 1
        
    ' Add found DUN to the combo box...
    If Result = ERROR_SUCCESS Then
      
      'get the connection name
      connName = Trim(connName)
      connName = Mid(connName, 1, Len(connName) - 1)
          
      'get profile of connection name
      tmpString = "RemoteAccess\Profile\" & connName
          
      'open
      If RegOpenKeyEx(HKEY_CURRENT_USER, tmpString, 0&, KEY_READ, KeyHandle2) = ERROR_SUCCESS Then
      
        'OK!
        tmpString = Space(MaxValue)
        
        'check all entries
        Do While RegEnumValue(KeyHandle2, 0, tmpString, MaxValue, 0&, REG_DWORD, ByVal 0&, MaxValue) = ERROR_SUCCESS
          
          'is it "User"
          'you can easily change this one to list all the Entries
          If InStr(1, tmpString, "user", vbTextCompare) > 0 Then
            'get the value for "User"
            RegQueryValueEx KeyHandle2, "user", &O0, REG_SZ, ByVal userName, MaxValue
            Exit Do
          End If
        Loop
      End If
          
      ' Close the registry key...
      RegCloseKey KeyHandle2
         
      ' Pad with spaces for a uniform look
      ' If your connName is larger, change MaxLength
      If Len(connName) < MaxLength Then
        connName = connName & Space(MaxLength - Len(connName))
      End If
         
      ' Add to Combo
      Combo1.AddItem connName & userName
    End If
  Loop While Result = ERROR_SUCCESS
   
  ' Close the registry key...
  RegCloseKey KeyHandle
  
  ' Enable Connect if we found a connection(s)
  Command1.Enabled = Combo1.ListCount > 0
  
  ' Start with first entry found
  Combo1.ListIndex = 0
End Sub


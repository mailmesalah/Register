VERSION 5.00
Begin VB.Form FRegister 
   Caption         =   "Register"
   ClientHeight    =   6270
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10590
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6270
   ScaleWidth      =   10590
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   885
      Left            =   825
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4335
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox TKey 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   480
      TabIndex        =   0
      Top             =   1665
      Width           =   9690
   End
   Begin VB.CommandButton CRegister 
      Caption         =   "Register"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   3855
      TabIndex        =   2
      Top             =   3645
      Width           =   2700
   End
   Begin VB.CommandButton CExtend 
      Caption         =   "Extend Registration"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   3855
      TabIndex        =   1
      Top             =   2655
      Width           =   2700
   End
   Begin VB.Label LSerialNo 
      Alignment       =   2  'Center
      Caption         =   "Serial No"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1800
      TabIndex        =   3
      Top             =   1035
      Width           =   7005
   End
End
Attribute VB_Name = "FRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim db As Database

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
        ByRef Destination As Any, ByRef Source As Any, ByVal numbytes As Long)

Private Const MAX_ADAPTER_NAME_LENGTH             As Long = 256
Private Const MAX_ADAPTER_DESCRIPTION_LENGTH      As Long = 128
Private Const MAX_ADAPTER_ADDRESS_LENGTH          As Long = 8
Private Const MAX_ADAPTER_NAME_LENGTH_p           As Long = MAX_ADAPTER_NAME_LENGTH + 4
Private Const MAX_ADAPTER_DESCRIPTION_LENGTH_p    As Long = MAX_ADAPTER_DESCRIPTION_LENGTH + 4
Private Const MIB_IF_TYPE_ETHERNET                As Long = 6

Private Type TIME_t
    aTime As Long
End Type

Private Type IP_ADDRESS_STRING
    IPadrString     As String * 16
End Type

Private Type IP_ADDR_STRING
    AdrNext         As Long
    IpAddress       As IP_ADDRESS_STRING
    IpMask          As IP_ADDRESS_STRING
    NTEcontext      As Long
End Type
' Information structure returned by GetIfEntry/GetIfTable
Private Type IP_ADAPTER_INFO
    Next As Long
    ComboIndex As Long
    AdapterName         As String * MAX_ADAPTER_NAME_LENGTH_p
    Description         As String * MAX_ADAPTER_DESCRIPTION_LENGTH_p
    MACadrLength        As Long
    MACaddress(0 To MAX_ADAPTER_ADDRESS_LENGTH - 1) As Byte
    AdapterIndex        As Long
    AdapterType         As Long             ' MSDN Docs say "UInt", but is 4 bytes
    DhcpEnabled         As Long             ' MSDN Docs say "UInt", but is 4 bytes
    CurrentIpAddress    As Long
    IpAddressList       As IP_ADDR_STRING
    GatewayList         As IP_ADDR_STRING
    DhcpServer          As IP_ADDR_STRING
    HaveWins            As Long             ' MSDN Docs say "Bool", but is 4 bytes
    PrimaryWinsServer   As IP_ADDR_STRING
    SecondaryWinsServer As IP_ADDR_STRING
    LeaseObtained       As TIME_t
    LeaseExpires        As TIME_t
End Type

Private Sub CExtend_Click()
Dim rs As Recordset, sKey As String
    
    Dim sMac As String

    sMac = GetMACs_AdaptInfo
    
    Dim sDate As String, r As Long, prvKey As String, sMacAd As String
    sDate = Format(Date & "", "DDMMMMYYYY")
        
    sMacAd = sMac
    sDate = Left(Base64EncodeString(sDate), 10)
    prvKey = Left(Base64EncodeString(sMacAd), 10)

    r = 1
    While r <= 10
        If ((r Mod 2) = 0) Then
            sKey = sKey & Mid(sDate, r, 1)
        Else
            sKey = sKey & Mid(prvKey, r, 1)
        End If
        r = r + 1
    Wend
    
    If UCase(sKey) = UCase(Trim(TKey.Text)) Then
    
        Set rs = db.OpenRecordset("Select * From Registration")
        While rs.EOF = False
            rs.Delete
            rs.MoveNext
        Wend
        
        rs.AddNew
        rs!RegDate = Date
        rs!RegKey = sKey
        rs.Update
        rs.Close
        
        MsgBox "Successfully Extended Demo Version."

    Else
        MsgBox "Wrong Key."
    End If
End Sub

Private Function GetDateExtendKey(sSerial As String) As String
    Dim sDate As String, r As Long, prvKey As String, sMacAd As String
    sDate = Format(Date & "", "DDMMMMYYYY")
        
    sMacAd = Decrement(Decrement(Mid(sSerial, 1, 1)))
    r = 2
    While r <= Len(sSerial)
        If ((r Mod 2) = 1) Then
            sMacAd = sMacAd & "-"
        End If
        
        sMacAd = sMacAd & Decrement(Decrement(Mid(sSerial, r, 1)))
        r = r + 1
    Wend
    
    sDate = Left(Base64EncodeString(sDate), 10)
    prvKey = Left(Base64EncodeString(sMacAd), 10)
    Dim sKey As String
    r = 1
    While r <= 10
        If ((r Mod 2) = 0) Then
            sKey = sKey & Mid(sDate, r, 1)
        Else
            sKey = sKey & Mid(prvKey, r, 1)
        End If
        r = r + 1
    Wend
    
    GetDateExtendKey = UCase(sKey)
End Function


Private Sub CRegister_Click()
    Dim macAd As String, pubKey As String, r As Long, sSerial As String

    macAd = GetMACs_AdaptInfo
    
    sSerial = GetSerial(macAd)
    LSerialNo.Caption = sSerial
    macAd = Left(Base64EncodeString(macAd), 10)
    pubKey = Left(Base64EncodeString("EC:A8:6B:FC:7E:A7"), 10)
    Dim sKey As String
    r = 1
    While r <= 10
        If ((r Mod 2) = 0) Then
            sKey = sKey & Mid(macAd, r, 1)
        Else
            sKey = sKey & Mid(pubKey, r, 1)
        End If
        r = r + 1
    Wend
    
    If (UCase(sKey) = Trim(TKey.Text)) Then
        Dim rs As Recordset
        'Register and make changes in Database
        Set rs = db.OpenRecordset("Select * From Registration")
        While rs.EOF = False
            rs.Delete
            rs.MoveNext
        Wend
        
        rs.AddNew
        rs!RegDate = Date
        rs!RegKey = sKey
        rs.Update
        rs.Close
        MsgBox "Successfully Registered."
    Else
        MsgBox "Wrong Key"
    End If
 
End Sub

Private Function GetSerial(sMac As String) As String
Dim r As Long, sSerial As String
    r = 1
    
    While r <= Len(sMac)
        If (Mid(sMac, r, 1) <> "-") Then
            sSerial = sSerial & Increment(Increment(Mid(sMac, r, 1)))
        End If
        
        r = r + 1
    Wend
    
    GetSerial = sSerial
End Function

Private Function Increment(sLet As String) As String
    
    If (UCase(sLet) = "A") Then
        Increment = "B"
    ElseIf (UCase(sLet) = "B") Then
        Increment = "C"
    ElseIf (UCase(sLet) = "C") Then
        Increment = "D"
    ElseIf (UCase(sLet) = "D") Then
        Increment = "E"
    ElseIf (UCase(sLet) = "E") Then
        Increment = "F"
    ElseIf (UCase(sLet) = "F") Then
        Increment = "G"
    ElseIf (UCase(sLet) = "G") Then
        Increment = "H"
    ElseIf (UCase(sLet) = "H") Then
        Increment = "I"
    ElseIf (UCase(sLet) = "I") Then
        Increment = "J"
    ElseIf (UCase(sLet) = "J") Then
        Increment = "K"
    ElseIf (UCase(sLet) = "K") Then
        Increment = "L"
    ElseIf (UCase(sLet) = "L") Then
        Increment = "M"
    ElseIf (UCase(sLet) = "M") Then
        Increment = "N"
    ElseIf (UCase(sLet) = "N") Then
        Increment = "O"
    ElseIf (UCase(sLet) = "O") Then
        Increment = "P"
    ElseIf (UCase(sLet) = "P") Then
        Increment = "Q"
    ElseIf (UCase(sLet) = "Q") Then
        Increment = "R"
    ElseIf (UCase(sLet) = "R") Then
        Increment = "S"
    ElseIf (UCase(sLet) = "S") Then
        Increment = "T"
    ElseIf (UCase(sLet) = "T") Then
        Increment = "U"
    ElseIf (UCase(sLet) = "U") Then
        Increment = "V"
    ElseIf (UCase(sLet) = "V") Then
        Increment = "W"
    ElseIf (UCase(sLet) = "W") Then
        Increment = "X"
    ElseIf (UCase(sLet) = "X") Then
        Increment = "Y"
    ElseIf (UCase(sLet) = "Y") Then
        Increment = "Z"
    ElseIf (UCase(sLet) = "Z") Then
        Increment = "0"
    ElseIf (UCase(sLet) = "0") Then
        Increment = "1"
    ElseIf (UCase(sLet) = "1") Then
        Increment = "2"
    ElseIf (UCase(sLet) = "2") Then
        Increment = "3"
    ElseIf (UCase(sLet) = "3") Then
        Increment = "4"
    ElseIf (UCase(sLet) = "4") Then
        Increment = "5"
    ElseIf (UCase(sLet) = "5") Then
        Increment = "6"
    ElseIf (UCase(sLet) = "6") Then
        Increment = "7"
    ElseIf (UCase(sLet) = "7") Then
        Increment = "8"
    ElseIf (UCase(sLet) = "8") Then
        Increment = "9"
    ElseIf (UCase(sLet) = "9") Then
        Increment = "A"
    Else
        Increment = ""
    End If
End Function

Private Function Decrement(sLet As String) As String
    
    If (UCase(sLet) = "A") Then
        Decrement = "9"
    ElseIf (UCase(sLet) = "B") Then
        Decrement = "A"
    ElseIf (UCase(sLet) = "C") Then
        Decrement = "B"
    ElseIf (UCase(sLet) = "D") Then
        Decrement = "C"
    ElseIf (UCase(sLet) = "E") Then
        Decrement = "D"
    ElseIf (UCase(sLet) = "F") Then
        Decrement = "E"
    ElseIf (UCase(sLet) = "G") Then
        Decrement = "F"
    ElseIf (UCase(sLet) = "H") Then
        Decrement = "G"
    ElseIf (UCase(sLet) = "I") Then
        Decrement = "H"
    ElseIf (UCase(sLet) = "J") Then
        Decrement = "J"
    ElseIf (UCase(sLet) = "K") Then
        Decrement = "J"
    ElseIf (UCase(sLet) = "L") Then
        Decrement = "K"
    ElseIf (UCase(sLet) = "M") Then
        Decrement = "L"
    ElseIf (UCase(sLet) = "N") Then
        Decrement = "M"
    ElseIf (UCase(sLet) = "O") Then
        Decrement = "N"
    ElseIf (UCase(sLet) = "P") Then
        Decrement = "O"
    ElseIf (UCase(sLet) = "Q") Then
        Decrement = "P"
    ElseIf (UCase(sLet) = "R") Then
        Decrement = "Q"
    ElseIf (UCase(sLet) = "S") Then
        Decrement = "R"
    ElseIf (UCase(sLet) = "T") Then
        Decrement = "S"
    ElseIf (UCase(sLet) = "U") Then
        Decrement = "T"
    ElseIf (UCase(sLet) = "V") Then
        Decrement = "U"
    ElseIf (UCase(sLet) = "W") Then
        Decrement = "V"
    ElseIf (UCase(sLet) = "X") Then
        Decrement = "W"
    ElseIf (UCase(sLet) = "Y") Then
        Decrement = "X"
    ElseIf (UCase(sLet) = "Z") Then
        Decrement = "Y"
    ElseIf (UCase(sLet) = "0") Then
        Decrement = "Z"
    ElseIf (UCase(sLet) = "1") Then
        Decrement = "0"
    ElseIf (UCase(sLet) = "2") Then
        Decrement = "1"
    ElseIf (UCase(sLet) = "3") Then
        Decrement = "2"
    ElseIf (UCase(sLet) = "4") Then
        Decrement = "3"
    ElseIf (UCase(sLet) = "5") Then
        Decrement = "4"
    ElseIf (UCase(sLet) = "6") Then
        Decrement = "5"
    ElseIf (UCase(sLet) = "7") Then
        Decrement = "6"
    ElseIf (UCase(sLet) = "8") Then
        Decrement = "7"
    ElseIf (UCase(sLet) = "9") Then
        Decrement = "8"
    Else
        Decrement = ""
    End If
End Function

Private Sub Form_Load()
    Set db = OpenDatabase(App.Path & "\Storage.mdb", False, False, "MS Access;PWD=12345abcde")
    
    LSerialNo.Caption = GetSerial(GetMACs_AdaptInfo())
End Sub

Public Function GetMACs_AdaptInfo() As String

    Dim AdapInfo As IP_ADAPTER_INFO, bufLen As Long, sts As Long
    Dim retStr As String, numStructs%, i%, IPinfoBuf() As Byte, srcPtr As Long
    
    
    ' Get size of buffer to allocate
    sts = GetAdaptersInfo(AdapInfo, bufLen)
    If (bufLen = 0) Then Exit Function
    numStructs = bufLen / Len(AdapInfo)
    
    ' reserve byte buffer & get it filled with adapter information
    ' !!! Don't Redim AdapInfo array of IP_ADAPTER_INFO,
    ' !!! because VB doesn't allocate it contiguous (padding/alignment)
    ReDim IPinfoBuf(0 To bufLen - 1) As Byte
    sts = GetAdaptersInfo(IPinfoBuf(0), bufLen)
    If (sts <> 0) Then Exit Function
    
    ' Copy IP_ADAPTER_INFO slices into UDT structure
    srcPtr = VarPtr(IPinfoBuf(0))
    For i = 0 To numStructs - 1
        If (srcPtr = 0) Then Exit For
'        CopyMemory AdapInfo, srcPtr, Len(AdapInfo)
        CopyMemory AdapInfo, ByVal srcPtr, Len(AdapInfo)
        
        ' Extract Ethernet MAC address
        With AdapInfo
            If (.AdapterType = MIB_IF_TYPE_ETHERNET) Then
                retStr = MAC2String(.MACaddress)
                Exit For
            End If
        End With
        srcPtr = AdapInfo.Next
    Next i
    
    ' Return list of MAC address(es)
    GetMACs_AdaptInfo = retStr
    
End Function

' Convert a zero-terminated fixed string to a dynamic VB string
Private Function sz2string(ByVal szStr As String) As String
    sz2string = Left$(szStr, InStr(1, szStr, Chr$(0)) - 1)
End Function

' Convert a byte array containing a MAC address to a hex string
Private Function MAC2String(AdrArray() As Byte) As String
    Dim aStr As String, hexStr As String, i%
    
    For i = 0 To 5
        If (i > UBound(AdrArray)) Then
            hexStr = "00"
        Else
            hexStr = Hex$(AdrArray(i))
        End If
        
        If (Len(hexStr) < 2) Then hexStr = "0" & hexStr
        aStr = aStr & hexStr
        If (i < 5) Then aStr = aStr & "-"
    Next i
    
    MAC2String = aStr
    
End Function


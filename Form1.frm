VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "Send SMS via Mobile Phone"
   ClientHeight    =   1575
   ClientLeft      =   12105
   ClientTop       =   2535
   ClientWidth     =   3795
   LinkTopic       =   "Form1"
   ScaleHeight     =   1575
   ScaleWidth      =   3795
   Begin MSCommLib.MSComm MSComm1 
      Left            =   2730
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send Now"
      Height          =   615
      Left            =   420
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Tested on Siemens(19200 baud) and Bosch (9600 baud)

Private Sub Command1_Click()
   Dim sms As String
   Dim buffer$
   
   MSComm1.CommPort = 2
   MSComm1.Settings = "19200,N,8,1"
   MSComm1.InputLen = 0
   MSComm1.Handshaking = comNone
   
   MSComm1.PortOpen = True
   
   'echo off
   MSComm1.Output = "atE0" & Chr$(13)
   Do
      DoEvents
   buffer$ = buffer$ & MSComm1.Input
   Loop Until InStr(buffer$, "OK")
   buffer$ = ""
   
   'set pdu mode
   MSComm1.Output = "AT+CMGF=0" & Chr$(13)
   Do
      DoEvents
   buffer$ = buffer$ & MSComm1.Input
   Loop Until InStr(buffer$, "OK")
   buffer$ = ""
   
   sms = MakeSms("004917212345678", "Does it work?")
   
   Debug.Print sms
   
   MSComm1.Output = "AT+CMGS=" & Len(sms) / 2 & Chr$(13)
   Do
      DoEvents
   buffer$ = buffer$ & MSComm1.Input
   Loop Until InStr(buffer$, ">")
   buffer$ = ""
   
   MSComm1.Output = sms + Chr$(26)
   
   Do
      DoEvents
   buffer$ = buffer$ & MSComm1.Input
   Loop Until InStr(buffer$, "OK")
   
   MSComm1.PortOpen = False
  
End Sub


Function MakeSms(number As String, txt As String) As String
  MakeSms = "001100"
  MakeSms = MakeSms + ConvNumber(number)
  MakeSms = MakeSms + "000064"
  MakeSms = MakeSms + ConvTxt(txt)
End Function


Function ConvNumber(num As String) As String

  Dim i As Integer
  Dim numType As String
  
  'default local number
  numType = "81"
  
  'but if international number then .....
  If Left$(num, 3) = "+00" And Len(num) > 3 Then num = Mid$(num, 4): numType = "91"
  If Left$(num, 2) = "00" And Len(num) > 2 Then num = Mid$(num, 3): numType = "91"
  If Left$(num, 1) = "+" And Len(num) > 1 Then num = Mid$(num, 2): numType = "91"
  
  
  ConvNumber = Right$("00" & Hex(Len(num)), 2)
  
  ConvNumber = ConvNumber + numType
  
  For i = 1 To Len(num) Step 2
    ConvNumber = ConvNumber + Mid$(num + "F", i + 1, 1) + Mid$(num + "F", i, 1)
  Next i
  
End Function


Function ConvTxt(txt As String) As String
  Dim i As Integer
  Dim datArr1(1 To 256) As Byte
  
  Dim l As Integer
  Dim touw As String
  
  
  'no more than 160 chars
  If Len(txt) > 160 Then txt = Left$(txt, 160)
  
  l = Len(txt)
  
  ConvTxt = Right$("00" & Hex(Len(txt)), 2)
  For i = 1 To l
    datArr1(i) = Asc(Mid$(txt, i, 1))
  Next i
  
  
  'make a bit stream of septets
  touw = ""
  For i = 1 To l
    touw = ToBin7(datArr1(i)) + touw
  Next i
  
  
  'and convert it to octets
  While Len(touw) > 8
    ConvTxt = ConvTxt + Bin2Hex(Right$(touw, 8))
    touw = Mid$(touw, 1, Len(touw) - 8)
  Wend
  
  ConvTxt = ConvTxt + Bin2Hex(touw)
  
  Debug.Print ConvTxt
End Function


Function ToBin7(ByVal num As Byte) As String
  'convert to padded 7 place binary number
  While num > 0
    ToBin7 = Trim(num Mod 2) + ToBin7
    num = num \ 2
  Wend
    
  ToBin7 = Right$("0000000" + ToBin7, 7)
End Function

Function Bin2Hex(ByVal touw As String) As String
  'convert binary to a padded 2 place hex number
  Dim x As Integer
  Dim num As Long
  
  For x = 1 To Len(touw)
    If Mid$(touw, x, 1) = "1" Then
      num = num + 2 ^ (Len(touw) - x)
    End If
  Next x
  
  Bin2Hex = Right$("00" + Hex(num), 2)
End Function


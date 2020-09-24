VERSION 5.00
Begin VB.Form fHex 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Convert To Hex"
   ClientHeight    =   1155
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manuell
   ScaleHeight     =   1155
   ScaleWidth      =   2025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Shape shpDrop 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   7  'Diagonalkreuz
      Height          =   615
      Left            =   1095
      Top             =   390
      Width           =   630
   End
   Begin VB.Label lb 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Drop .BIN file here:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   1665
   End
End
Attribute VB_Name = "fHex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

  Dim hFile     As Long
  Dim HexName   As String
  Dim BinName   As String
  Dim i         As Long
  Dim b()       As Byte
  Dim hx        As String

    hFile = FreeFile
    On Error Resume Next
        Open Data.Files(1) For Binary As hFile
        If Err Then
            MsgBox "Files only, please", vbExclamation
          ElseIf LCase$(Right$(Data.Files(1), 4)) <> ".bin" Then   'ERR = FALSE/0
            MsgBox "Only .BIN files, please", vbExclamation
          Else 'NOT LCASE$(RIGHT$(DATA.FILES(1),...
            i = InStrRev(Data.Files(1), "\")
            BinName = Replace$(Mid$(Data.Files(1), i + 1), ".", "_")
            HexName = Left$(BinName, Len(BinName) - 3) & "hex"
            ReDim b(1 To LOF(1))
            Get #hFile, , b
            Close hFile
            For i = 1 To UBound(b)
                hx = hx & Right$("0" & Hex$(b(i)), 2) & " "
                If (i Mod 80) = 0 And i <> UBound(b) Then 'insert linebreak
                    hx = hx & "|"
                End If
            Next i
            hx = Replace$(hx, "|", """ & _" & vbCrLf & Space$(38) & """")
            Clipboard.SetText "" & _
                              "Private Declare Sub MemCopy Lib ""kernel32"" Alias ""RtlMoveMemory"" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)" & vbCrLf & vbCrLf & _
                              "Private Const " & HexName & " As String = """ & RTrim$(hx) & """" & vbCrLf & _
                              "Private " & BinName & "() As Byte " & vbCrLf & vbCrLf & _
                              "Private Sub Patch(ByVal FON As Long, HexCode As String, BinCode() As Byte) 'FON is the Function's Ordinal Number in vTable" & vbCrLf & vbCrLf & _
                              "  'Convert hex to binary and patch vTable entry for Function xxx" & vbCrLf & vbCrLf & _
                              "  Dim s()           As String" & vbCrLf & _
                              "  Dim p             As Long" & vbCrLf & _
                              "  Dim CodeAddress   As Long" & vbCrLf & _
                              "  Dim VTableAddress As Long" & vbCrLf & vbCrLf & _
                              "    'Concert hex to binary" & vbCrLf & _
                              "    s = Split(HexCode, "" "")" & vbCrLf & _
                              "    Redim Bincode(0 To UBound(s))" & vbCrLf & _
                              "    For p = 0 To UBound(s)" & vbCrLf & _
                              "        BinCode(p) = Val(""&H"" & s(p))" & vbCrLf & _
                              "    Next p" & vbCrLf & vbCrLf & _
                              "    'Patch Function xxx vTable entry" & vbCrLf & _
                              "    CodeAddress = VarPtr(BinCode(0))" & vbCrLf & _
                              "    MemCopy VarPtr(VTableAddress), ObjPtr(Me), 4 'get vTable address" & vbCrLf & _
                              "    MemCopy VTableAddress + FON * 4 + 28, VarPtr(CodeAddress), 4 'patch proper entry in vTable" & vbCrLf & vbCrLf & _
                              "End Sub"
            MsgBox "Paste clipboard contents into your code...", vbInformation
            Unload Me
        End If
    On Error GoTo 0

End Sub

':) Ulli's VB Code Formatter V2.21.6 (2006-Apr-07 15:17)  Decl: 1  Code: 61  Total: 62 Lines
':) CommentOnly: 0 (0%)  Commented: 3 (4,8%)  Empty: 4 (6,5%)  Max Logic Depth: 4

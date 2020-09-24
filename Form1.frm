VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RichTx32.ocx"
Begin VB.Form Form1 
   Caption         =   "Tool Bar"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   9315
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox tooltip 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   600
      ScaleHeight     =   255
      ScaleWidth      =   8295
      TabIndex        =   1
      Top             =   1680
      Visible         =   0   'False
      Width           =   8295
   End
   Begin RichTextLib.RichTextBox RTF 
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   8493
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      TextRTF         =   $"Form1.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Const EM_LINEINDEX = &HBB
Const EM_GETFIRSTVISIBLELINE = &HCE

Dim funName() As String '// Collecting the function name
Dim funIn() As String   '// Collecting the inside of the function
Dim c As Integer

Private Sub Form_Load()

    Dim getData As String
    Dim strLine() As String
    Dim spl() As String
    Dim i As Integer
    
    tooltip.BackColor = RGB(255, 255, 225)
    
    getData = GetFileData(App.Path & "\function.txt")
    strLine = Split(getData, vbCrLf)

    For i = LBound(strLine) To UBound(strLine)
        spl = Split(strLine(i), ":")
        AddData spl(0), spl(1)
    Next i

End Sub

Private Sub Form_Resize()

    On Error Resume Next

    RTF.Width = Me.Width - 125
    RTF.Height = Me.Height - 400

End Sub

Private Sub RTF_Change()

    helpTool

End Sub

Private Sub RTF_SelChange()

    helpTool

End Sub

Private Sub helpTool()

    Dim lStart As Long      '// line start
    Dim lEnd As Long        '// line end
    Dim tStart As Integer   '// the sel start of strLine
    Dim strLine As String   '// get the current line text
    
    Dim intOpen As Integer  '// ( open
    Dim intClose As Integer '// ) close
    
    Dim functionName As String '// Get the function name
    Dim oneChr As String
    
    Dim strDelimiter As String  '// block the text
    
    Dim i As Integer        '// Counter
    
    Dim StrIn As String     '// String inside the (...)
    
    '// Set Delimiter
    strDelimiter = ",(){}[]-+*%/= '~!&|<>?:;.#" & Chr(34) & vbTab
    
    '// Get start and end position
    lStart = SendMessage(RTF.hwnd, EM_LINEINDEX, RTF.GetLineFromChar(RTF.SelStart), 0&)
    lEnd = SendMessage(RTF.hwnd, EM_LINEINDEX, RTF.GetLineFromChar(RTF.SelStart) + 1, 0&)

    If lEnd <= 0 Then lEnd = Len(RTF.Text)

    '// Get current line text and selstart
    strLine = Mid(RTF.Text, lStart + 1, lEnd)
    tStart = RTF.SelStart - lStart
    
    If tStart <= 0 Then tStart = 1
    
    '// Check if it in (...)
    intOpen = InStrRev(strLine, "(", tStart)
    intClose = InStrRev(strLine, ")", tStart)
    
    If intOpen > intClose Then
        '// It in ()

        '// find function name
        For i = intOpen - 1 To 1 Step -1
            oneChr = Mid(strLine, i, 1)
            If InStr(1, strDelimiter, oneChr) > 0 Then
                Exit For
            Else
                functionName = oneChr & functionName
            End If
        Next i
        
        '// Setting the position
        tooltip.Left = tooltip.TextWidth(Left(strLine, intOpen - (1 + Len(functionName))))
        tooltip.Top = ((RTF.GetLineFromChar(RTF.SelStart) - _
                    SendMessage(RTF.hwnd, EM_GETFIRSTVISIBLELINE, 0&, 0&)) + 1.5) * _
                    tooltip.TextHeight("ABC")

        '// get text in (...)
        intClose = InStr(tStart, strLine, ")")
        If intClose <= 0 Then intClose = Len(strLine)
        
        StrIn = Mid(strLine, intOpen + 1, intClose - intOpen)
        
        PrintTool functionName, StrIn, UBound(Split(Mid(strLine, intOpen, tStart - intOpen + 1), ","))

    Else
        '// Not in ()
        tooltip.Visible = False
    End If

End Sub

Private Sub PrintTool(functionName As String, StrIn As String, reqNum As Integer)

    Dim fp As Integer   '// function position
    Dim reqFun() As String  '// requirment of function
    Dim reqHave() As String '// requirment that we have put inside
    
    '// Check if function is exist or not
    fp = isinList(functionName)
    If fp <= 0 Then Exit Sub
    
    '// Show Tooltip
    tooltip.Visible = True
    tooltip.Cls
    
    '// Set width
    tooltip.Height = 255
    tooltip.Width = tooltip.TextWidth(" " & funName(fp) & "(" & funIn(fp) & ")")
    If InStr(1, funIn(fp), vbCrLf) Then
        tooltip.Height = 255 * 2
        tooltip.Width = tooltip.TextWidth(Left(funIn(fp), 75))
    End If
    
    '// Draw Border
    tooltip.Line (0, 0)-(tooltip.Width, 0), RGB(212, 208, 200)
    tooltip.Line (0, 0)-(0, tooltip.Height), RGB(212, 208, 200)
    tooltip.Line (0, tooltip.Height - 10)-(tooltip.Width, tooltip.Height - 10), RGB(64, 64, 64)
    tooltip.Line (tooltip.Width - 10, 0)-(tooltip.Width - 10, tooltip.Height), RGB(64, 64, 64)
    
    '// Set position X and Y
    tooltip.CurrentX = 0
    tooltip.CurrentY = 0
    
    '// Print function name
    tooltip.ForeColor = vbBlack
    tooltip.Print " " & funName(fp) & "(";
    
    '// bold the current require of function
    reqFun = Split(funIn(fp), ",")
    reqHave = Split(StrIn, ",")
    
    If UBound(reqFun) < reqNum Then
        tooltip.Print funIn(fp);
    Else
        For i = LBound(reqFun) To UBound(reqFun)
            If i = reqNum Then
                tooltip.FontBold = True
                tooltip.Print reqFun(i);
                tooltip.FontBold = False
            Else
                tooltip.Print reqFun(i);
            End If
            
            If i <> UBound(reqFun) Then
                tooltip.Print ",";
            End If
        Next i
    End If
    
    tooltip.Print ") ";

End Sub

Private Function isinList(name As String) As Integer

    '// Check if it exist in list of not
    '// if it exist it will return FunctionName position
    '// if not return Nothing
    
    Dim i As Integer
    
    For i = LBound(funName) To UBound(funName)
        If UCase(funName(i)) = UCase(name) Then
            isinList = i
            Exit Function
            Exit For
        End If
    Next i

End Function

Private Sub AddData(str1 As String, str2 As String)

    c = c + 1

    ReDim Preserve funName(c) As String
    ReDim Preserve funIn(c) As String

    funName(c) = str1
    If Len(str2) > 75 Then
        str2 = Left(str2, 75) & vbCrLf & Right(str2, Len(str2) - 75)
    End If
    If str2 = " " Then str2 = ""
    funIn(c) = str2

End Sub

Function GetFileData(strFile As String) As String

Open strFile For Input As #1
    GetFileData = StrConv(InputB(LOF(1), 1), vbUnicode)
Close #1

End Function

VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Equation Drawer"
   ClientHeight    =   9450
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11190
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   630
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   746
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Clear"
      Height          =   255
      Left            =   15
      TabIndex        =   28
      Top             =   8160
      Width           =   495
   End
   Begin VB.CommandButton cmdInt 
      Caption         =   "Do"
      Height          =   255
      Left            =   9000
      TabIndex        =   26
      Top             =   9120
      Width           =   375
   End
   Begin VB.TextBox intTo 
      Height          =   285
      Left            =   8280
      TabIndex        =   25
      Text            =   "0"
      Top             =   9120
      Width           =   735
   End
   Begin VB.TextBox intFrom 
      Height          =   285
      Left            =   7320
      TabIndex        =   23
      Text            =   "0"
      Top             =   9120
      Width           =   735
   End
   Begin VB.CommandButton cmdDrawTangent 
      Caption         =   "Draw Tangent"
      Height          =   255
      Left            =   8040
      TabIndex        =   20
      Top             =   8760
      Width           =   1335
   End
   Begin VB.TextBox dydx 
      Height          =   285
      Left            =   6960
      TabIndex        =   19
      Text            =   "0"
      Top             =   8760
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Help !"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3960
      TabIndex        =   16
      Top             =   8520
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "PreScales"
      Height          =   255
      Left            =   9960
      TabIndex        =   15
      Top             =   8400
      Width           =   1095
   End
   Begin VB.TextBox increaSE 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4560
      TabIndex        =   14
      Text            =   "1"
      Top             =   9060
      Width           =   735
   End
   Begin VB.TextBox ToX 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   12
      Text            =   "50"
      Top             =   9000
      Width           =   615
   End
   Begin VB.TextBox FromX 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   960
      TabIndex        =   9
      Text            =   "-50"
      Top             =   9000
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Scale"
      Height          =   255
      Left            =   9360
      TabIndex        =   8
      Top             =   8400
      Width           =   615
   End
   Begin VB.TextBox Xscale 
      Height          =   285
      Left            =   8400
      TabIndex        =   6
      Text            =   "6"
      Top             =   8400
      Width           =   975
   End
   Begin VB.TextBox Yscale 
      Height          =   285
      Left            =   6600
      TabIndex        =   4
      Text            =   "6"
      Top             =   8400
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "DRAW"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3120
      TabIndex        =   2
      Top             =   8520
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   480
      TabIndex        =   1
      Text            =   "x"
      Top             =   8520
      Width           =   2655
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      DrawStyle       =   6  'Inside Solid
      DrawWidth       =   2
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8295
      Left            =   0
      ScaleHeight     =   549
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   741
      TabIndex        =   0
      Top             =   0
      Width           =   11175
      Begin VB.Label EqLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Label6"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   0
         Left            =   -600
         TabIndex        =   17
         Top             =   240
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Line LY 
         BorderColor     =   &H00E0E0E0&
         Index           =   0
         X1              =   0
         X2              =   808
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line LX 
         BorderColor     =   &H00E0E0E0&
         Index           =   0
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   552
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   744
         Y1              =   272
         Y2              =   272
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         X1              =   368
         X2              =   368
         Y1              =   0
         Y2              =   680
      End
   End
   Begin VB.Label lblArea 
      Caption         =   "AREA"
      Height          =   255
      Left            =   9480
      TabIndex        =   27
      Top             =   9120
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "To"
      Height          =   255
      Index           =   2
      Left            =   8040
      TabIndex        =   24
      Top             =   9120
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "Integration from X="
      Height          =   255
      Index           =   1
      Left            =   5880
      TabIndex        =   22
      Top             =   9120
      Width           =   1575
   End
   Begin VB.Label lblGradient 
      Caption         =   "GRADIENT"
      Height          =   255
      Left            =   9480
      TabIndex        =   21
      Top             =   8760
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "dy/dx AT X="
      Height          =   255
      Index           =   0
      Left            =   5880
      TabIndex        =   18
      Top             =   8760
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "X-Step="
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3720
      TabIndex        =   13
      Top             =   9120
      Width           =   735
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   " > X >"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   11
      Top             =   9120
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Domain:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   9120
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "X Scale"
      Height          =   255
      Left            =   7680
      TabIndex        =   7
      Top             =   8400
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Y Scale"
      Height          =   255
      Left            =   5880
      TabIndex        =   5
      Top             =   8400
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Y="
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   8520
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E8CDB9&
      BorderColor     =   &H00E8CDB9&
      FillColor       =   &H00E8CDB9&
      FillStyle       =   0  'Solid
      Height          =   1455
      Left            =   0
      Top             =   8040
      Width           =   5535
   End
   Begin VB.Menu PreScale 
      Caption         =   "PreScale"
      Visible         =   0   'False
      Begin VB.Menu sincostan 
         Caption         =   "Sin,Cos,Tan"
      End
      Begin VB.Menu hyper 
         Caption         =   "Hyperbolic Functions"
      End
      Begin VB.Menu qube 
         Caption         =   "Qubic Equations"
      End
      Begin VB.Menu logar 
         Caption         =   "Log(x)"
      End
      Begin VB.Menu ceq 
         Caption         =   "Complex Equations"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text


Private Sub SetColor(ByVal EquationID As Integer)

While EquationID > 5
EquationID = EquationID - 5
Wend


Select Case EquationID
Case 0
Picture1.ForeColor = vbBlack
Case 1
Picture1.ForeColor = 128
Case 2
Picture1.ForeColor = &H1D7917
Case 3
Picture1.ForeColor = &H3D4354
Case 4
Picture1.ForeColor = &H632769
Case 5
Picture1.ForeColor = &HF2974D
End Select
End Sub

Private Sub Proccessing(ByVal IsProccessing As Boolean)
IsProccessing = Not IsProccessing
FromX.Enabled = IsProccessing ' Disable All Buttons and TextBoxes To Avoid Errors if They Were Changed During Drawing
ToX.Enabled = IsProccessing
increaSE.Enabled = IsProccessing
Command1.Enabled = IsProccessing
Command2.Enabled = IsProccessing
Command3.Enabled = IsProccessing
Command4.Enabled = IsProccessing
Command5.Enabled = IsProccessing
Text1.Enabled = IsProccessing
Yscale.Enabled = IsProccessing
Xscale.Enabled = IsProccessing
dydx.Enabled = IsProccessing
cmdDrawTangent.Enabled = IsProccessing
intFrom.Enabled = IsProccessing
intTo.Enabled = IsProccessing
cmdInt.Enabled = IsProccessing

End Sub

Private Function Xcordinate(P As String) As String ' To Convert From X Cordinate To Pixels
Xcordinate = 368 + P * Xscale
End Function

Private Function Ycordinate(P As String) As String ' To Convert From Y Cordinate To Pixels
Ycordinate = 272 - P * Yscale

End Function

Private Sub ceq_Click()
Xscale = 30
Yscale = 30
Command2_Click
increaSE = 0.1
End Sub

Private Sub cmdDrawTangent_Click()
On Error Resume Next
Dim i2 As Integer, dcs() As String, P1 As String, P2 As String
lblGradient = ""
Dim gradient As Double
Dim Y_Intercept As Double


dcs = Split(Text1, ",") ' To see if more than one equation exist

For i2 = 0 To UBound(dcs) ' pass equation one by one
SetColor i2
    P1 = GetVal(ResolveValue(Me.dydx - 0.0001, dcs(i2)))
    P2 = GetVal(ResolveValue(Me.dydx - -0.0001, dcs(i2)))

'Y=M*X+C   where M is gradient
'(dydx-0.0001 , P1) , (dydx+0.0001 , P2)

gradient = (P1 - P2) / ((dydx - 0.0001) - (dydx - -0.0001))
gradient = Round(gradient, 5)
lblGradient = lblGradient & gradient & ", "
'C (Y-intercept)=Y-M*X
Y_Intercept = P1 - gradient * (dydx - 0.0001)
Y_Intercept = Round(Y_Intercept, 5)
Picture1.Line (Xcordinate(-80), Ycordinate((-80 * gradient) + Y_Intercept))-(Xcordinate(80), Ycordinate((80 * gradient) + Y_Intercept))
'MsgBox Round(gradient, 5), , Y_Intercept

Next i2
lblGradient = Mid(lblGradient, 1, Len(lblGradient) - 2)
End Sub

Private Sub cmdInt_Click()
On Error Resume Next
'Using Strings To Hold Numbers May Be A Slower Proccess, But More Accurate
Dim i2 As Integer, dcs() As String, P1 As String, P2 As String, Area As String
lblArea = ""
Proccessing True
Dim gradient As Double
Dim Y_Intercept As Double


dcs = Split(Text1, ",") ' To see if more than one equation exist

For i2 = 0 To UBound(dcs) ' pass equation one by one
SetColor i2
Text1.Tag = GetVal(ResolveValue(i, dcs(i2)))
Area = 0
For i = intFrom To intTo - 0.005 Step 0.005 ' this loop is done to draw the curve

DoEvents

Me.Caption = " Progress = " & Round((((i - intFrom) * 100) / (intTo - intFrom)), 0) & "%"
   DoEvents
    Me.Tag = Text1.Tag ' GetVal(ResolveValue(i, dcs(i2)))
   DoEvents

    Text1.Tag = GetVal(ResolveValue(i + 0.005, dcs(i2)))
    Area = Area - -(((Me.Tag - -Text1.Tag) / 2) * 0.005)
'lblArea = Area
DoEvents
   Picture1.Line (Xcordinate(Str(i)), Ycordinate(Me.Tag))-(Xcordinate(i + 0.005), Ycordinate(0))
   DoEvents

Next i
Area = Round(Val(Area), 1)
lblArea = lblArea & Area & ", "
EqLabel(i2 + 1).ZOrder 1    'To Bring The Label On Top Of The Drawing

EqLabel(i2 + 1).ZOrder 0

Next i2

lblArea = Mid(lblArea, 1, Len(lblArea) - 2) ' To Remove The ", " on the right !
Proccessing False
End Sub

Private Sub Command1_Click()
On Error Resume Next



Proccessing True ' Disable All Buttons and TextBoxes To Avoid Errors if They Were Changed During Drawing

Picture1.Cls
Picture1.DrawMode = 13
Picture1.ForeColor = vbBlack
Dim i As Double, dcs() As String, i2 As Double, MostSuitableXForLable As Long, MostSuitableYForLable As Long

For i = 1 To EqLabel.UBound
Unload EqLabel(i)
Next i

dcs = Split(Text1, ",") ' To see if more than one equation exist

For i2 = 0 To UBound(dcs) ' pass equation one by one
Text1.Tag = GetVal(ResolveValue(FromX, dcs(i2)))
DoEvents
SetColor i2
Load EqLabel(i2 + 1)
EqLabel(i2 + 1).Caption = " Y=" & dcs(i2) & " "

    For i = FromX To ToX - increaSE Step increaSE ' this loop is done to draw the curve

DoEvents

Me.Caption = " Progress = " & Round((((i - FromX) * 100) / (ToX - increaSE - FromX)), 0) & "%"
   DoEvents
    Me.Tag = Text1.Tag ' GetVal(ResolveValue(i, dcs(i2)))
   DoEvents


    Text1.Tag = GetVal(ResolveValue(i + increaSE, dcs(i2)))
DoEvents
If Xcordinate(Str(i)) >= 0 And Xcordinate(Str(i)) < Picture1.Width - EqLabel(i2 + 1).Width And Ycordinate(Me.Tag) >= 0 And Ycordinate(Me.Tag) <= Picture1.Height - EqLabel(0).Height Then MostSuitableXForLable = Xcordinate(Str(i)): MostSuitableYForLable = Ycordinate(Me.Tag)
    Picture1.Line (Xcordinate(Str(i)), Ycordinate(Me.Tag))-(Xcordinate(i + increaSE), Ycordinate(Text1.Tag))
   DoEvents

    Next i
' to label the Curve
    ''Picture1.ForeColor = &HE2A35C
    ''Picture1.Line (MostSuitableXForLable, MostSuitableYForLable)-(MostSuitableXForLable, MostSuitableYForLable)
    ''Picture1.Print "Y=" & dcs(i2)
    EqLabel(i2 + 1).BackColor = Picture1.ForeColor
    EqLabel(i2 + 1).Left = MostSuitableXForLable
    EqLabel(i2 + 1).Top = MostSuitableYForLable
    EqLabel(i2 + 1).Visible = True
    EqLabel(i2 + 1).ZOrder 0
    Me.Caption = " Progress = 100%"
    
    
DoEvents
Next i2


Proccessing False
Exit Sub
exitsubnow:
Exit Sub

End Sub

Private Function GetVal(eq As String) As String
On Error GoTo exitsun
DoEvents

GetVal = StringCalc(eq)

Exit Function
exitsun:
DoEvents

Exit Function

End Function


Private Function ResolveValue(ByVal iVal As Double, ByVal Equation As String) As String

Dim ss() As String, ii As Integer
While InStr(Equation, "x") > 0
Equation = MyReplace(Equation, "x", iVal)
Wend
ResolveValue = Equation


End Function


Private Function MyReplace(ByVal Expr As String, ByVal FindVal As String, ByVal ReplaceVal As String) As String

MyReplace = Mid(Expr, 1, InStr(Expr, FindVal) - 1) & ReplaceVal & Mid(Expr, InStr(Expr, FindVal) + 1)


End Function

Private Sub Command2_Click()
Dim i As Integer
For i = 1 To EqLabel.UBound
Unload EqLabel(i)
Next i



Picture1.Cls

For i = 1 To 124
DoEvents
LX(i).X1 = Xcordinate(62 - i)
LX(i).X2 = LX(i).X1
DoEvents
Next i
For i = 1 To 92
DoEvents
LY(i).Y1 = Ycordinate(46 - i)
LY(i).Y2 = LY(i).Y1
DoEvents
Next i
'MsgBox Round((Picture1.Width) / Xscale, 1)
If Val(FromX) < Round((-368) / Xscale, 1) Then FromX = Round((-368) / Xscale, 1)
If Val(ToX) > Round((368) / Xscale, 1) Then ToX = Round((368) / Xscale, 1)

End Sub



Private Sub Command3_Click()

PopupMenu Me.PreScale
End Sub



Private Sub Command4_Click()
Form4.Show
End Sub

Private Sub Command5_Click()
Picture1.Cls
Dim i As Integer
For i = 1 To EqLabel.UBound
Unload EqLabel(i)
Next i
End Sub

Private Sub Form_Load()
Me.Caption = "This Program Was Done By Sami Haija --> ExpLoSioN System Series"

Dim i As Integer
For i = 1 To 124
Load LX(i)
LX(i).X1 = Xcordinate(62 - i)
LX(i).X2 = LX(i).X1
LX(i).Visible = True
Next i
For i = 1 To 92
Load LY(i)
LY(i).Y1 = Ycordinate(46 - i)
LY(i).Y2 = LY(i).Y1
LY(i).Visible = True
Next i
Picture1.DrawMode = 13

End Sub




Private Function StringCalc(ByVal CalcWhat As String) As Double

' This Function was written by sami haija .. U MUST EMAIL SAMIHAIJA@YAHOO.COM IN ORDER TO USE IT !



While CalcWhat Like "*--*" 'To Replace the -- by a +
'MsgBox "Entering"
CalcWhat = Mid(CalcWhat, 1, InStr(CalcWhat, "--") - 1) & "+" & Mid(CalcWhat, InStr(CalcWhat, "--") + 2)
'MsgBox CalcWhat
Wend

Dim iii As Integer, FirstOperation As String

If Len(CalcWhat) = 0 Then StringCalc = 0: Exit Function

Dim Ff() As String


If Not CalcWhat Like "*(*" Then ' If the equation contains ( )
If IsNumeric(CalcWhat) Then StringCalc = CalcWhat: Exit Function ' no need to enter the function if the number was already numeric

    For iii = Len(CalcWhat) To 2 Step -1 ' from Right To Left Since the Last Done Operation have the Priority
    
    If Mid(CalcWhat, iii, 1) = "-" And IsNumeric(Mid(CalcWhat, iii - 1, 1)) = True Then
    FirstOperation = "-"
    Exit For
    End If
    
    If Mid(CalcWhat, iii, 1) = "+" And IsNumeric(Mid(CalcWhat, iii - 1, 1)) = True Then
    FirstOperation = "+"
    Exit For
    End If
    
    Next iii
        
        ' Search for -,+ and then for *,/
        ' When a - or + is found
        ' all the string at the left of the - is send to the same function and it's right too and substracted from each other
        ' Ex: 2*2-20  --> stringcalc(2*2) - stringcalc(20)
        ' so the * will be automatically performed first
        ' got it ??

        If FirstOperation = "+" Then
        StringCalc = StringCalc(Mid(CalcWhat, 1, iii - 1)) + StringCalc(Mid(CalcWhat, iii + 1))
        Exit Function
        End If
        
        If FirstOperation = "-" Then
        StringCalc = StringCalc(Mid(CalcWhat, 1, iii - 1)) - StringCalc(Mid(CalcWhat, iii + 1))
        Exit Function
        End If
        
        ' and now search for / , *

    For iii = Len(CalcWhat) To 2 Step -1
    
    If Mid(CalcWhat, iii, 1) = "/" And IsNumeric(Mid(CalcWhat, iii - 1, 1)) = True Then
    FirstOperation = "/"
    Exit For
    End If
    If Mid(CalcWhat, iii, 1) = "*" And IsNumeric(Mid(CalcWhat, iii - 1, 1)) = True Then
    FirstOperation = "*"
    Exit For
    End If
    Next iii
        
        If FirstOperation = "*" Then
        StringCalc = StringCalc(Mid(CalcWhat, 1, iii - 1)) * StringCalc(Mid(CalcWhat, iii + 1))
        Exit Function
        End If


        If FirstOperation = "/" Then
        StringCalc = StringCalc(Mid(CalcWhat, 1, iii - 1)) / StringCalc(Mid(CalcWhat, iii + 1))
        Exit Function
        End If



Ff = Split(CalcWhat, "sin") ' to see if a sine exists
If UBound(Ff) > 0 Then
StringCalc = Sin(Ff(1))
Exit Function
End If


Ff = Split(CalcWhat, "cos")  ' to see if a cosine exists
If UBound(Ff) > 0 Then
StringCalc = Cos(Ff(1))
Exit Function
End If


Ff = Split(CalcWhat, "tan") ' to see if a tan exists
If UBound(Ff) > 0 Then
StringCalc = Tan(Ff(1))
Exit Function
End If


Ff = Split(CalcWhat, "log") ' to see if a logarithm exists
If UBound(Ff) > 0 Then
StringCalc = Log((Ff(1)))
Exit Function
End If

Ff = Split(CalcWhat, "^") ' the first priority ... the power !
If UBound(Ff) > 0 Then
StringCalc = ((StringCalc(Ff(0)) ^ StringCalc(Mid(CalcWhat, Len(Ff(0)) + 2))))
Exit Function
End If


Else ' this is done when (brackets) are found

Dim ii As Integer, cc As Integer, i As Integer, Soution As Double
ii = InStr(CalcWhat, "(")
cc = 1

For i = ii + 1 To Len(CalcWhat) - i  ' this loop detects Nested Brackets
If Mid(CalcWhat, i, 1) = "(" Then cc = cc + 1
If Mid(CalcWhat, i, 1) = ")" Then cc = cc - 1



If cc = 0 Then
Soution = StringCalc(Mid(CalcWhat, ii + 1, i - ii - 1))
StringCalc = StringCalc(Mid(CalcWhat, 1, ii - 1) & Soution & Mid(CalcWhat, i + 1))
Exit Function
End If
Next i


End If


End Function

Private Sub hyper_Click()
Xscale = 45
increaSE = 0.1
Yscale = 45
Command2_Click
End Sub

Private Sub logar_Click()
Xscale = 25
Yscale = 40
Command2_Click
increaSE = 0.1
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1.ToolTipText = "(" & Round((X - 368) / Xscale, 1) & "," & Round((272 - Y) / Yscale, 1) & ")"
'Picture1.ToolTipText = X & " " & Y
End Sub

Private Sub qube_Click()
Xscale = 40
Yscale = 25
Command2_Click
increaSE = 0.1
End Sub

Private Sub sincostan_Click()
Xscale = 40
increaSE = 0.09
Yscale = 40
Command2_Click
End Sub

Private Sub Text1_LostFocus()
While InStr(Text1, " ") > 0
Text1 = Mid(Text1, 1, InStr(Text1, " ") - 1) & Mid(Text1, InStr(Text1, " ") + 1)
Wend
End Sub

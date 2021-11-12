VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Demo - motif detection"
   ClientHeight    =   11265
   ClientLeft      =   5880
   ClientTop       =   1950
   ClientWidth     =   17580
   LinkTopic       =   "Form1"
   Picture         =   "WM.frx":0000
   ScaleHeight     =   751
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1172
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox y_axis 
      Caption         =   "Reverse axis"
      Height          =   350
      Left            =   9840
      TabIndex        =   46
      Top             =   0
      Width           =   1215
   End
   Begin VB.Frame MatrixXYZ 
      Caption         =   "General information"
      Height          =   5055
      Left            =   4920
      TabIndex        =   9
      Top             =   5280
      Width           =   7215
      Begin VB.CommandButton ShowVariants 
         Caption         =   "Text matrix"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton ShowVariants 
         Caption         =   "Count matrix"
         Height          =   375
         Index           =   1
         Left            =   1560
         TabIndex        =   26
         ToolTipText     =   "Position Frequency Matrix (PFM)"
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton ShowVariants 
         Caption         =   "Weight matrix"
         Height          =   375
         Index           =   2
         Left            =   2880
         TabIndex        =   25
         ToolTipText     =   "Position Probability Matrix (PPM)"
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton ShowVariants 
         Caption         =   "LogRatio matrix"
         Height          =   375
         Index           =   3
         Left            =   4200
         TabIndex        =   24
         ToolTipText     =   " Position Weight Matrix (PWM)"
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton ShowVariants 
         Caption         =   "Info"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   5760
         TabIndex        =   23
         Top             =   360
         Width           =   1215
      End
      Begin VB.PictureBox IF_holder 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         ForeColor       =   &H80000008&
         Height          =   4215
         Left            =   240
         ScaleHeight     =   279
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   447
         TabIndex        =   22
         Top             =   720
         Width           =   6735
         Begin VB.PictureBox NPSet 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   1575
            Left            =   5040
            Picture         =   "WM.frx":0446
            ScaleHeight     =   103
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   95
            TabIndex        =   42
            Top             =   120
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.PictureBox FreqPanel 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   2175
            Left            =   240
            ScaleHeight     =   143
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   415
            TabIndex        =   34
            Top             =   1800
            Width           =   6255
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00C0C0C0&
            X1              =   16
            X2              =   416
            Y1              =   64
            Y2              =   64
         End
         Begin VB.Label w_pos 
            BackStyle       =   0  'Transparent
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Left            =   240
            TabIndex        =   43
            Top             =   600
            Width           =   6255
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Unordered Sequence Logo:"
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Left            =   240
            TabIndex        =   36
            Top             =   1440
            Width           =   6255
         End
      End
      Begin VB.PictureBox MP_holder 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   4215
         Left            =   240
         ScaleHeight     =   277
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   445
         TabIndex        =   20
         Top             =   720
         Visible         =   0   'False
         Width           =   6735
         Begin VB.CheckBox Fallow_M 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fallow Matrix !"
            Height          =   375
            Left            =   600
            TabIndex        =   32
            Top             =   3120
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.TextBox Mlog 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            Height          =   285
            Index           =   0
            Left            =   5760
            TabIndex        =   21
            Text            =   "0"
            Top             =   120
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VB.PictureBox MW_holder 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   4215
         Left            =   240
         ScaleHeight     =   277
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   445
         TabIndex        =   18
         Top             =   720
         Visible         =   0   'False
         Width           =   6735
         Begin VB.CheckBox FollowWM 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Follow Weight Matrix"
            Height          =   255
            Left            =   480
            TabIndex        =   33
            Top             =   3120
            Value           =   1  'Checked
            Width           =   2055
         End
         Begin VB.TextBox Mwght 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Index           =   0
            Left            =   5760
            TabIndex        =   19
            Text            =   "0"
            Top             =   120
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VB.PictureBox MV_holder 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   4215
         Left            =   240
         ScaleHeight     =   277
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   445
         TabIndex        =   16
         Top             =   720
         Visible         =   0   'False
         Width           =   6735
         Begin VB.TextBox Mval 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Height          =   285
            Index           =   0
            Left            =   5760
            TabIndex        =   17
            Text            =   "0"
            Top             =   120
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VB.PictureBox MT_holder 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   4215
         Left            =   240
         ScaleHeight     =   277
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   445
         TabIndex        =   14
         Top             =   720
         Visible         =   0   'False
         Width           =   6735
         Begin VB.TextBox MTxt 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
            Height          =   285
            Index           =   0
            Left            =   5760
            TabIndex        =   15
            Text            =   "0"
            Top             =   120
            Visible         =   0   'False
            Width           =   495
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Parameters"
      Height          =   5055
      Left            =   12360
      TabIndex        =   6
      Top             =   5280
      Width           =   5055
      Begin VB.Frame Frame5 
         Caption         =   "Make score values from:"
         Height          =   2415
         Left            =   240
         TabIndex        =   28
         Top             =   2400
         Width           =   4575
         Begin VB.CheckBox LR_DANU 
            Caption         =   "Use the values from Position Weight Matrix (PWM)"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   31
            ToolTipText     =   "sum from the LogRatio matrix ..."
            Top             =   480
            Value           =   1  'Checked
            Width           =   4215
         End
         Begin VB.CommandButton Scann 
            Caption         =   "Scan for motif"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   240
            TabIndex        =   30
            Top             =   1560
            Width           =   4095
         End
         Begin VB.CheckBox LR_DANU 
            Caption         =   "Use the values from Position Probability Matrix (PPM)"
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   29
            ToolTipText     =   "sum from the Weight matrix ..."
            Top             =   960
            Width           =   4215
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Zoom on signal"
         Height          =   1455
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   4575
         Begin VB.OptionButton ZSO 
            Caption         =   "Zoom signal"
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   360
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton SLO 
            Caption         =   "Signal level (with follow matrix option)"
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   840
            Width           =   3975
         End
         Begin VB.HScrollBar Zoom_Scale 
            Height          =   255
            Left            =   1560
            Max             =   200
            Min             =   4
            TabIndex        =   11
            Top             =   360
            Value           =   30
            Width           =   2895
         End
      End
      Begin VB.Timer RealTime 
         Interval        =   100
         Left            =   4440
         Top             =   240
      End
      Begin VB.TextBox Lungime_fereastra 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3000
         TabIndex        =   7
         Text            =   "30"
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sliding window length:"
         Height          =   255
         Left            =   960
         TabIndex        =   8
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "The motif set (input)"
      Height          =   5055
      Left            =   240
      TabIndex        =   3
      Top             =   5280
      Width           =   4455
      Begin VB.OptionButton Val_Z 
         Caption         =   "Initialize PFM with zero values"
         Height          =   255
         Left            =   240
         TabIndex        =   45
         ToolTipText     =   "Position Frequency Matrix (PFM), also known as the count matrix"
         Top             =   3840
         Value           =   -1  'True
         Width           =   3975
      End
      Begin VB.OptionButton Val_P 
         Caption         =   "Initialize PFM with pseudocounts (0.000001)"
         Height          =   255
         Left            =   240
         TabIndex        =   44
         ToolTipText     =   "Position Frequency Matrix (PFM), also known as the count matrix"
         Top             =   3480
         Width           =   3975
      End
      Begin VB.TextBox WMatrix 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Text            =   "WM.frx":2C78
         Top             =   360
         Width           =   3975
      End
      Begin VB.CommandButton Make_Matrix 
         Caption         =   "Make Position weight matrix (PWM)"
         Height          =   615
         Left            =   240
         TabIndex        =   4
         Top             =   4200
         Width           =   3975
      End
   End
   Begin VB.PictureBox Center_patt 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   4695
      Left            =   12360
      ScaleHeight     =   309
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   333
      TabIndex        =   2
      Top             =   360
      Width           =   5055
      Begin VB.Line Line10 
         BorderStyle     =   3  'Dot
         X1              =   168
         X2              =   168
         Y1              =   0
         Y2              =   312
      End
      Begin VB.Line Line9 
         BorderStyle     =   3  'Dot
         X1              =   336
         X2              =   -8
         Y1              =   152
         Y2              =   152
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   240
      ScaleHeight     =   157
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   717
      TabIndex        =   1
      Top             =   360
      Width           =   10815
      Begin VB.Label down_label 
         BackStyle       =   0  'Transparent
         Caption         =   "Different from the motif set (red)"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   2040
         Width           =   5175
      End
      Begin VB.Label up_label 
         BackStyle       =   0  'Transparent
         Caption         =   "Like the motif set (blue)"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   120
         Width           =   4455
      End
      Begin VB.Line Zero 
         BorderColor     =   &H00404040&
         X1              =   0
         X2              =   728
         Y1              =   80
         Y2              =   80
      End
      Begin VB.Shape Window_Shape 
         BorderColor     =   &H00808080&
         Height          =   4095
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   15
      End
   End
   Begin VB.TextBox secventata 
      Height          =   2055
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "WM.frx":2CE7
      Top             =   3120
      Width           =   11895
   End
   Begin VB.Line Line6 
      X1              =   744
      X2              =   744
      Y1              =   176
      Y2              =   192
   End
   Begin VB.Line Line5 
      X1              =   744
      X2              =   744
      Y1              =   16
      Y2              =   32
   End
   Begin VB.Line Line4 
      X1              =   744
      X2              =   720
      Y1              =   24
      Y2              =   24
   End
   Begin VB.Line Line3 
      X1              =   744
      X2              =   720
      Y1              =   184
      Y2              =   184
   End
   Begin VB.Label down_val 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11280
      TabIndex        =   48
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label up_val 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11280
      TabIndex        =   47
      Top             =   240
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   -7680
      Picture         =   "WM.frx":343C
      Top             =   10560
      Width           =   25290
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Zoom into the signal using the sliding window length:"
      Height          =   255
      Left            =   12360
      TabIndex        =   41
      Top             =   120
      Width           =   4815
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "The main signal over the entire sequence:"
      Height          =   255
      Left            =   240
      TabIndex        =   40
      Top             =   120
      Width           =   5415
   End
   Begin VB.Label Label4 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11280
      TabIndex        =   37
      Top             =   1440
      Width           =   135
   End
   Begin VB.Line Line1 
      X1              =   728
      X2              =   744
      Y1              =   104
      Y2              =   104
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Find motif in sequence:"
      Height          =   255
      Left            =   240
      TabIndex        =   35
      Top             =   2880
      Width           =   4935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   ________________________________                          ____________________
'  /    Demo - motif detection      \________________________/       v1.00        |
' |                                                                               |
' |            Name:  Demo - motif detection V1.0                                 |
' |        Category:  Open source software                                        |
' |          Author:  Paul A. Gagniuc                                             |
' |            Book:  Algorithms in Bioinformatics: Theory and Implementation     |
' |           Email:  paul_gagniuc@acad.ro                                        |
' |  ____________________________________________________________________________ |
' |                                                                               |
' |    Date Created:  January 2014                                                |
' |          Update:  August 2021                                                 |
' |       Tested On:  WinXP, WinVista, Win7, Win8, Win10                          |
' |             Use:  Analysis of splice sites                                    |
' |                                                                               |
' |                  _____________________________                                |
' |_________________/                             \_______________________________|
'

Dim M() As String           'stores the rows for Matrix ()
Dim Weight() As String      'proportions
Dim Matrix() As String      'matrix of nucleotides
Dim Matrix_Val() As String  'no of A and T and C and G in cols of Matrix()
Dim Weight_Max() As String  'order by max values
Dim MLogRatio() As String   'Log weight

'######################
Dim MX_pos As Variant
Dim MY_pos As Variant

Dim space_x As Integer
Dim space_y As Integer
'######################
Dim Interface_on As Boolean
Dim select_start_reminder As Integer


Private Sub Form_Load()

    Interface_on = False

    '##################
    MX_pos = 20
    MY_pos = 25
    
    space_x = 3
    space_y = 3
    '##################

    secventata.Text = Replace(secventata.Text, Chr(10), "")
    secventata.Text = Replace(secventata.Text, Chr(13), "")
    
    'X
    Line9.X1 = 0
    Line9.X2 = Center_patt.ScaleWidth
    Line9.Y1 = (Center_patt.ScaleHeight / 2)
    Line9.Y2 = (Center_patt.ScaleHeight / 2)
    Line9.Visible = True
    
    'Y
    Line10.X1 = (Center_patt.ScaleWidth / 2)
    Line10.X2 = (Center_patt.ScaleWidth / 2)
    Line10.Y1 = 0
    Line10.Y2 = Center_patt.ScaleWidth
    Line10.Visible = True
    
    Line1.Y1 = Picture1.Top + (Picture1.ScaleHeight / 2) + 2
    Line1.Y2 = Line1.Y1
    
    Zero.Y1 = (Picture1.ScaleHeight / 2)
    Zero.Y2 = Zero.Y1
    
    Label4.Top = Line1.Y1 - (Label4.Height / 2) + 3
    
    'Make_Matrix_Click
    
    DoEvents
    Scann_Click

End Sub

Private Sub LR_DANU_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    For i = 0 To LR_DANU.UBound
        LR_DANU(i).Value = 0
    Next i
    
    LR_DANU(Index).Value = 1
    
    Scann_Click
    DoEvents
    ShowVariants_Click (4)
    
End Sub

'##########################################################################
'##########################################################################
'##########################################################################

Private Sub Make_Matrix_Click()
    
    'Format textbox in case of bad input
    '--------------------------------------------------------------------------
    If Len(WMatrix.Text) < 3 Then Exit Sub
    
    WMatrix.Text = UCase(WMatrix.Text)
    WMatrix.Text = Replace(WMatrix.Text, vbCrLf & vbCrLf, vbCrLf)
    
    Dim ttmp() As String
    ttmp() = Split(WMatrix.Text, vbCrLf)
    If (ttmp(UBound(ttmp())) > "") Then WMatrix.Text = WMatrix.Text & vbCrLf
    
    s = WMatrix.Text
    t = (Len(s) - Len(Replace(s, vbCrLf, ""))) / 2 '[vbcrlf represents two chr]
    Frame1.Caption = "The motif set (input): " & t & " motifs"
    
    '--------------------------------------------------------------------------

    M() = Split(WMatrix.Text, vbCrLf)
    Rows = UBound(M()) - 1
    cols = Len(M(0)) - 1
    
    '------ Matrix(n,n) umplut cu caractere -----------------------------------
    
    ReDim Matrix(0 To Rows, 0 To cols) As String
    
    For i = 0 To Rows
    
        For j = 0 To cols
        
            Matrix(i, j) = Mid(M(i), j + 1, 1)
            'MsgBox "Matrix(" & i & "|" & j & ")" & Mid(M(i), j + 1, 1)
    
        Next j
    
    Next i
    
    '--------------------------------------------------------------------------

    '------ Matrix_Val(n,4) umplut cu valori ----------------------------------
    
    Dim Xb(0 To 3) As String
    ReDim Matrix_Val(0 To 3, 0 To cols) As String
    
    
    For j = 0 To cols
        
        For e = 0 To 3
            If Val_Z.Value = True Then Xb(e) = 0
            If Val_P.Value = True Then Xb(e) = 0.000001
        Next e
        
        
        For i = 0 To Rows
        
            n = Matrix(i, j)
            If n = "A" Then Xb(0) = Xb(0) + 1
            If n = "C" Then Xb(1) = Xb(1) + 1
            If n = "G" Then Xb(2) = Xb(2) + 1
            If n = "T" Then Xb(3) = Xb(3) + 1
    
        Next i
        
    
        For v = 0 To 3
            Matrix_Val(v, j) = Xb(v)
            'MsgBox "col = " & j & vbCrLf & "v= " & Xb(v)
        Next v
    
    Next j
    
    '--------------------------------------------------------------------------
        
    '------ Weight Matrix (0 to 1 values for each nucleotide in the row) ------
    ReDim Weight(0 To cols, 0 To 3) As String
    Dim nx(0 To 3) As String
    
    For j = 0 To cols
        
        For v = 0 To 3
            nx(v) = Val(Matrix_Val(v, j))
        Next v
        
        For v = 0 To 3
        
            'Weight(j, v) = (1 / (Val(nx(0)) + Val(nx(1)) + Val(nx(2)) + Val(nx(3)))) * Val(Matrix_Val(v, j))

            If Val_Z.Value = True Then Weight(j, v) = Round(Val(Matrix_Val(v, j)) / (Rows + 1), 1)
            If Val_P.Value = True Then Weight(j, v) = Val(Matrix_Val(v, j)) / (Rows + 1)

            '----------------------
            'If v = 0 Then X = "A"
            'If v = 1 Then X = "C"
            'If v = 2 Then X = "G"
            'If v = 3 Then X = "T"

            'MsgBox "col = " & j & vbCrLf & "v= " & X & vbCrLf & Matrix_Val(v, j) & vbCrLf & Weight(j, v)
            '---------------------
        Next v
            
    Next j
    
    '--------------------------------------------------------------------------

    '------ MLogRatio (Log ratio = N/0.25) ------------------------------------
    ReDim MLogRatio(0 To cols, 0 To 3) As String
    Dim nl(0 To 3) As String
    
    For j = 0 To cols
        
        For v = 0 To 3
        
            If Val(Weight(j, v)) = 0 Then
                MLogRatio(j, v) = 0
            Else
                MLogRatio(j, v) = Round(Log(Val(Weight(j, v)) / 0.25), 3)
            End If
            
            '----------------------
            'If v = 0 Then X = "A"
            'If v = 1 Then X = "C"
            'If v = 2 Then X = "G"
            'If v = 3 Then X = "T"
    
            'MsgBox "col = " & j & vbCrLf & "v= " & X & vbCrLf & vbCrLf & MLogRatio(j, v)
            '---------------------
        Next v
            
    Next j
    
    '--------------------------------------------------------------------------
    
    Call FreqNuc(Len(M(0)))

    ShowVariants_Click (4)
    
End Sub


Function FreqNuc(ByVal cols As Integer)

    Dim aY(0 To 3) As Integer
    xBuc = FreqPanel.ScaleWidth / cols
    ybuc = FreqPanel.ScaleHeight / 100
    
    For i = 1 To cols
    
        plus = 0
    
        For j = 0 To 3
    
            If j = 0 Then
                sX = 0
                sY = 0
                sW = 46
                sH = 47
                xColor = vbGreen   'A
            End If
            
            If j = 1 Then
                sX = 46
                sY = 0
                sW = 41
                sH = 47
                xColor = vbRed     'C
            End If
            
            If j = 2 Then
                sX = 43
                sY = 49
                sW = 44
                sH = 49
                xColor = vbBlue    'G
            End If
            
            If j = 3 Then
                sX = 4
                sY = 50
                sW = 38
                sH = 47
                xColor = vbYellow  'T
            End If
    
    
            aY(j) = Val(Weight(i - 1, j)) * 100 'Nx4 (Weight(0 To Cols, 0 To 3))
            'FreqPanel.Line (xBuc * oldi, plus)-(xBuc * i, plus + (ybuc * aY(j))), xColor, BF
            
            
            pX = xBuc * oldi
            pY = plus
            
            pWith = xBuc
            pHeight = (ybuc * aY(j)) + 1
            
            FreqPanel.PaintPicture NPSet.Picture, pX, pY, pWith, pHeight, sX, sY, sW, sH, vbSrcCopy
            
            plus = plus + (ybuc * aY(j))
            'Debug.Print aY(j)
            
        Next j
    
        oldi = i
    
    Next i

    
End Function

'Public Function SortArray(ByRef TheArray As Variant)
    'Sorted = False
    'Do While Not Sorted
    '    Sorted = True
    'For X = 0 To UBound(TheArray) - 1
    '    If TheArray(X) > TheArray(X + 1) Then
    '        Temp = TheArray(X + 1)
    '        TheArray(X + 1) = TheArray(X)
    '        TheArray(X) = Temp
    '        Sorted = False
    '    End If
    'Next X
    'Loop
'End Function


Private Sub Swap(a As Variant, b As Variant)
    Dim t As Variant
    'Swap a and b
    t = a
    a = b
    b = t
End Sub
 
Private Sub Sort(arr As Variant)

    Dim X As Long, Y As Long, Size As Long
    On Error Resume Next
     
        'Get array size
        Size = UBound(arr)
        
        'Do the sorting.
        For X = 0 To Size
            For Y = X To Size
                If arr(Y) < arr(X) Then
                    'Swap vals
                    Call Swap(arr(X), arr(Y))
                End If
            Next Y
        Next X
        
End Sub


Function FreqMaxNuc(ByVal cols As Integer)

    ReDim Weight_Max(1 To cols) As String
    
    Dim aY(0 To 3) As Integer
    
    
    For c = 1 To cols
    
        For i = 0 To 3
            For j = 0 To 3
        
                aY(j) = Val(Weight(i - 1, j) * 100) 'Nx4 (Weight(0 To Cols, 0 To 3))
                
            Next j
        Next i
    
    Next c

End Function


Function Ask_matrix(ByVal n As String, ByVal j As Integer) As String

        ' ask frequency of nucleotide n
        If n = "A" Then v = 0
        If n = "C" Then v = 1
        If n = "G" Then v = 2
        If n = "T" Then v = 3
        
        If LR_DANU(0).Value = 1 Then
            Ask_matrix = MLogRatio(j, v)
        End If
        
        If LR_DANU(1).Value = 1 Then
            Ask_matrix = Weight(j, v)
        End If
        
End Function


Function Read_matrix(ByVal winstr As String) As String
    'read window
    
    If LR_DANU(0).Value = 1 Then
        '--------------------------------------------------------------------------
        If SLO.Value = True And Fallow_M.Value = 1 Then Call Light_Erase(Mlog, &HC0FFC0)  'Follow Log Matrix
        If SLO.Value = True And FollowWM.Value = 1 Then Call Light_Erase(Mwght, &HC0FFFF) 'Follow Weight Matrix
        
        For j = 1 To Len(winstr)
        
            nuc = Mid(winstr, j, 1)
        
            'Filter, it will show only the blue signal ...
            
            'If Val(Ask_matrix(nuc, j - 1)) = 0 Then
            '    q = 0
            '    j = Len(winstr)
            'Else
            '    q = Val(q) + Val(Ask_matrix(nuc, j - 1)) 'nucleotide and the col position in the matrix
            'End If
        
            If Val(q) = 0 Then
                q = Val(Ask_matrix(nuc, j - 1))
            Else
                q = Val(q) + Val(Ask_matrix(nuc, j - 1)) 'nucleotide and the col position in the matrix
            End If
            
            If SLO.Value = True And Fallow_M.Value = 1 Then Call Light_TextBox(j - 1, nuc, Mlog)  'Follow Log Matrix
            If SLO.Value = True And FollowWM.Value = 1 Then Call Light_TextBox(j - 1, nuc, Mwght) 'Follow Weight Matrix
            
        Next j
        '--------------------------------------------------------------------------
    End If
    
    
    
    If LR_DANU(1).Value = 1 Then
        '--------------------------------------------------------------------------
        If SLO.Value = True And Fallow_M.Value = 1 Then Call Light_Erase(Mlog, &HC0FFC0)  'Follow Log Matrix
        If SLO.Value = True And FollowWM.Value = 1 Then Call Light_Erase(Mwght, &HC0FFFF) 'Follow Weight Matrix
        
        For j = 1 To Len(winstr)
        
            nuc = Mid(winstr, j, 1)
        
            'Filter, it will show only the blue signal ...
            
            If Val(Ask_matrix(nuc, j - 1)) = 0 Then
                q = 0
                j = Len(winstr)
            Else
                q = Val(q) + Val(Ask_matrix(nuc, j - 1)) 'nucleotide and the col position in the matrix
            End If
            
            If SLO.Value = True And Fallow_M.Value = 1 Then Call Light_TextBox(j - 1, nuc, Mlog)  'Follow Log Matrix
            If SLO.Value = True And FollowWM.Value = 1 Then Call Light_TextBox(j - 1, nuc, Mwght) 'Follow Weight Matrix
            
        Next j
        '--------------------------------------------------------------------------
    End If
    
    
    Read_matrix = q
End Function


Function Light_Erase(ByRef Obj As Object, ByVal color As Double)

    Dim c, r, a, i As Integer
    
    a = 0
    
    Y = 4
    X = Len(M(0)) '- 1
    
    
    For r = 1 To Y
    
        For c = 1 To X
    
            a = a + 1
            Obj(a).BackColor = color
            'Obj(a).BackColor = &HC0FFC0
            'Obj(a).BackColor = &HC0FFFF
            'Obj(a).Refresh
        Next c
    
    Next r
End Function


Function Light_TextBox(ByVal j As Integer, ByVal nuc As String, ByRef Obj As Object)

    Dim c, r, a, i As Integer
    
    Y = 4
    X = Len(M(0)) '- 1
    
    a = 0
    
    For r = 1 To Y
        
        For c = 1 To X
    
            a = a + 1
    
            If nuc = "A" Then v = 1
            If nuc = "C" Then v = 2
            If nuc = "G" Then v = 3
            If nuc = "T" Then v = 4
    
            If j = c - 1 And v = r Then
                Obj(a).BackColor = &H808080
                'Obj(a).Refresh
            End If
    
        Next c
    
    Next r

End Function


'##########################################################################
'##########################################################################
'##########################################################################

Function Fill_AllTypeOfMatrix(ByVal X As Integer, ByVal Y As Integer, ByVal Matrix_type As String, ByRef Obj As Object, ByRef PicObj As Object)

    Dim c, r, a, i, MXX As Integer
    Dim xx, yy As Variant
    
    For i = 1 To Obj.UBound
        Unload Obj(i)
    Next i
    
    
    a = 0
    
    For r = 1 To Y
        
        For c = 1 To X
    
            a = a + 1
    
            Load Obj(a)
            
            xx = MX_pos + ((Obj(0).Width + space_x) * c)
            yy = MY_pos + ((Obj(0).Height + space_y) * r)
    
            Obj(a).Move xx, yy
            Obj(a).Visible = True
            
            If Matrix_type = 1 Then Obj(a).Text = Matrix(r - 1, c - 1)     'NxN (Matrix(0 To Rows, 0 To Cols))
            If Matrix_type = 2 Then Obj(a).Text = Matrix_Val(r - 1, c - 1) '4xN (Matrix_Val(0 To 3, 0 To Cols))
            If Matrix_type = 3 Then Obj(a).Text = Weight(c - 1, r - 1)     'Nx4 (Weight(0 To Cols, 0 To 3))
            If Matrix_type = 4 Then Obj(a).Text = MLogRatio(c - 1, r - 1)  'Nx4 (MLogRatio(0 To Cols, 0 To 3))
            
            Obj(a).Refresh
    
        Next c
    
    Next r
    
    'Paint numbers and lines on any matrix
    '#######################
    PicObj.Cls
    PicObj.Line (Obj(1).Left - 20, Obj(1).Top - 10)-(Obj(1).Left + ((Obj(1).Width + space_x) * X) + 20, Obj(1).Top - 10), vbBlack
    PicObj.Line (Obj(1).Left - 10, Obj(1).Top - 20)-(Obj(1).Left - 10, Obj(1).Top + ((Obj(1).Height + space_y) * Y) + 15), vbBlack
    
    For i = 0 To X - 1
        PicObj.CurrentX = (Obj(1).Left + (Obj(1).Width + space_x) * i) + 5
        PicObj.CurrentY = Obj(1).Top - 35
        PicObj.Font.Size = 16
        PicObj.Print i + 1
    Next i
    
    If Matrix_type <> 1 Then
        For i = 0 To 3
    
            If i = 0 Then s = "A"
            If i = 1 Then s = "C"
            If i = 2 Then s = "G"
            If i = 3 Then s = "T"
    
            PicObj.CurrentX = Obj(1).Left - 30
            PicObj.CurrentY = (Obj(1).Top + (Obj(1).Height + space_y) * i)
            PicObj.Font.Size = 15
            PicObj.Print s
        
        Next i
    End If
    '#######################

End Function


'##########################################################################
'##########################################################################
'##########################################################################

Function Fill_All()

    'Show Matrix of Nucleotides
    '#######################################
    Rows = UBound(M()) '- 1
    cols = Len(M(0)) '- 1
    
    If Rows < 12 Then
        Call Fill_AllTypeOfMatrix(cols, Rows, 1, MTxt, MT_holder)
    Else
        MsgBox "Rows are biger than 12. Thus, the Test matrix from the" & _
        "general info panel will not be ploted for a visual inspection!"
    End If
    '#######################################
    
    'Show Matrix Count.
    '#######################################
    Rows = 4
    cols = Len(M(0)) '- 1
    Call Fill_AllTypeOfMatrix(cols, Rows, 2, Mval, MV_holder)
    '#######################################
    
    'Show Matrix Prob.
    '#######################################
    Rows = 4
    cols = Len(M(0)) '- 1
    Call Fill_AllTypeOfMatrix(cols, Rows, 3, Mwght, MW_holder)
    '#######################################
    
    'Show Matrix LogRatio
    '#######################################
    Rows = 4
    cols = Len(M(0)) '- 1
    Call Fill_AllTypeOfMatrix(cols, Rows, 4, Mlog, MP_holder)
    '#######################################

End Function



Function k_real_time()

    Dim y_scale As Variant

    X = select_start_reminder 'global
    
    If ZSO.Value = True Then
    '----------------------------------------------------
        If Interface_on = False Then Exit Function
        
        cols = Len(M(0)) '- 1 ' window size = motif size
        Lungime_fereastra.Text = Zoom_Scale.Value 'Cols + Zoom_Scale.Value
        
        '---------------------
        q = (Len(secventata.Text) / Picture1.ScaleWidth) * X
        If q < 0 Then Exit Function
        
        secventata.SetFocus
        secventata.SelStart = q
        
        secventata.SelLength = Lungime_fereastra.Text
        secventaADN = secventata.SelText
        
        w_pos.Caption = "Sliding window from " & Int(q) & "b to " & Int(q) + Val(Lungime_fereastra.Text) & "b"
    
        If Len(secventata.Text) < 3 Then Exit Function
    
        Window_Shape.Width = (Picture1.ScaleWidth / Len(secventata.Text)) * Len(secventaADN)
        Window_Shape.Left = select_start_reminder
        '---------------------
        
        If Len(secventaADN) <= 2 Then Exit Function
        If (Len(secventaADN) - cols) <= 0 Then Exit Function
        
        y_scale = Val(up_val.Caption)
        
        buc = Center_patt.ScaleWidth / (Len(secventaADN) - cols)
        bucH = Center_patt.ScaleHeight / (y_scale * 2)
    
        Center_patt.Cls
        
        For h = 1 To Len(secventaADN) - cols
    
            Window = Mid(secventaADN, h, cols)
            val_motif = Val(Read_matrix(Window))
                
            OLD_val_motif = val_motif
            
            If OLD_val_motif <> 0 Then
                If val_motif < 0 Then
                    Center_patt.Line (buc * oldh, (Center_patt.ScaleHeight / 2))-(buc * h, (Center_patt.ScaleHeight / 2) - (bucH * val_motif)), vbRed, BF
                Else
                    Center_patt.Line (buc * oldh, (Center_patt.ScaleHeight / 2))-(buc * h, (Center_patt.ScaleHeight / 2) - (bucH * val_motif)), vbBlue, BF
                End If
            End If
    
    
            'OLD_val_motif = val_motif
            'If OLD_val_motif <> 0 Then
            '   Center_patt.Line (buc * oldh, (Center_patt.ScaleHeight / 2))-(buc * h, (Center_patt.ScaleHeight / 2) - (bucH * val_motif)), vbRed, BF
            'End If
    
            'If OLD_val_motif = 0 Then
            '   Center_patt.Line (buc * oldh, (Center_patt.ScaleHeight / 2) - (bucH * OLD_val_motif))-(buc * h, (Center_patt.ScaleHeight / 2) - (bucH * val_motif)), vbRed ', BF
            'Else
            '   Center_patt.Line (buc * oldh, (Center_patt.ScaleHeight / 2) - (bucH * OLD_val_motif))-(buc * h, (Center_patt.ScaleHeight / 2) - (bucH * val_motif)), vbRed ', BF
            'End If
            'OLD_val_motif = val_motif
    
    
            'OLD_val_motif = val_motif
            'If OLD_val_motif <> 0 Then
            '   Center_patt.Line (buc * oldh, 0)-(buc * h, (Center_patt.ScaleHeight / 2) - (bucH * val_motif)), vbRed, BF
            'End If
    
    
            'OLD_val_motif = val_motif
            'If OLD_val_motif <> 0 Then
            '   Center_patt.Line (buc * oldh, (Center_patt.ScaleHeight / 2) - (bucH * OLD_val_motif))-(buc * h, (Center_patt.ScaleHeight / 2) - (bucH * val_motif)), vbRed ', BF
            'End If
    
            oldh = h
                
    
        Next h
    '----------------------------------------------------
    End If
    
    
    
    If SLO.Value = True Then
    '----------------------------------------------------
        cols = Len(M(0)) '- 1 ' window size = motif size
        Lungime_fereastra.Text = cols
    
    
        '---------------------
        q = (Len(secventata.Text) / Picture1.ScaleWidth) * X
    
        secventata.SetFocus
        secventata.SelStart = q
        secventata.SelLength = Lungime_fereastra.Text
        secventaADN = secventata.SelText
    
        w_pos.Caption = "Sliding window from " & Int(q) & "b to " & Int(q) + Val(Lungime_fereastra.Text) & "b"
    
        If Len(secventata.Text) < 3 Then Exit Function
    
        Window_Shape.Width = (Picture1.ScaleWidth / Len(secventata.Text)) * Len(secventaADN)
        Window_Shape.Left = X
        '---------------------
    
        If Len(secventaADN) < 2 Then Exit Function
        
        y_scale = Val(up_val.Caption)
        
        buc = Center_patt.ScaleWidth / (Len(secventaADN))
        bucH = Center_patt.ScaleHeight / (y_scale * 2)
    
        Center_patt.Cls
        
        For h = 1 To buc
    
            val_motif = Val(Read_matrix(secventaADN))
    
                Center_patt.Line (buc * oldh, (Center_patt.ScaleHeight / 2) - (bucH * OLD_val_motif))-(buc * h, (Center_patt.ScaleHeight / 2) - (bucH * val_motif)), vbBlack
    
                Center_patt.CurrentX = 10
                Center_patt.CurrentY = 10
                Center_patt.Font.Size = 16
                Center_patt.Print "Log ratio = " & val_motif
    
                oldh = h
                OLD_val_motif = val_motif
    
        Next h
    '----------------------------------------------------
    End If

End Function


Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Window_Shape.Visible = True
    select_start_reminder = X
    
    Call k_real_time

End Sub

Private Sub RealTime_Timer()
    If ZSO.Value = True Then
        If Lungime_fereastra.Text <> Zoom_Scale.Value Then
            Lungime_fereastra.Text = Zoom_Scale.Value
            Call k_real_time
        End If
    End If
End Sub

Private Sub Scann_Click()

    Dim Maxyp As Variant
    Dim Maxym As Variant
    Dim y_scale As Variant
    
    Interface_on = True
    'RealTime.Enabled = False
    
    secventata.Text = Replace(secventata.Text, Chr(10), "")
    secventata.Text = Replace(secventata.Text, Chr(13), "")
    
    
    Picture1.Cls
    '--------------------------------------------------------------------------
    Make_Matrix_Click
    
    DoEvents
    
    Call Fill_All
    '--------------------------------------------------------------------------
    ss = secventata.Text
    cols = Len(M(0)) '- 1 ' window size = motif size
    
    
    Maxyp = 0
    Maxym = 0
    
    For i = 1 To Len(ss) - cols
    
        Window = Mid(ss, i, cols)
        val_motif = Val(Read_matrix(Window))
        
        If val_motif > Maxyp Then Maxyp = val_motif
        If val_motif < Maxym Then Maxym = val_motif
        
    Next i
    
    
    If Maxyp = 0 Then Maxyp = 1
    If Maxym = 0 Then Maxym = -1
    

    'This reverses the scale (+) down and (-) up
    If y_axis.Value = 1 Then
    
        If Maxyp > Abs(Maxym) Then y_scale = -(Maxyp) Else y_scale = Maxym
        
        up_val.Caption = Round(y_scale, 2)
        down_val.Caption = Round(-y_scale, 2)
        
        up_label.Caption = "Different from the motif set (red)"
        down_label.Caption = "Like the motif set (blue)"
    
    End If
    
    'This shows normal scale, (+) up and (-) down
    If y_axis.Value = 0 Then
    
        If Maxyp > Abs(Maxym) Then y_scale = Maxyp Else y_scale = Abs(Maxym)
        
        up_val.Caption = Round(y_scale, 2)
        down_val.Caption = Round(-y_scale, 2)
        
        up_label.Caption = "Like the motif set (blue)"
        down_label.Caption = "Different from the motif set (red)"
    
    End If
    
    
    sliceX = (Picture1.ScaleWidth / (Len(ss) - cols))
    sliceY = (Picture1.ScaleHeight / (y_scale * 2))

    sliceYIC = (Picture1.ScaleHeight / 100)
    
    Zero.X1 = 0
    Zero.X2 = Picture1.ScaleWidth
    Zero.Y1 = (Picture1.ScaleHeight / 2)
    Zero.Y2 = (Picture1.ScaleHeight / 2)
    Zero.Visible = True
    
    '--------------------------------------------------------------------------

    For i = 1 To Len(ss) - cols
    
        Window = Mid(ss, i, cols)
        val_motif = Read_matrix(Window)
    
        OLD_val_motif = val_motif
        
        If OLD_val_motif <> 0 Then
            If val_motif < 0 Then
                Picture1.Line (sliceX * oldi, (Picture1.ScaleHeight / 2))-(sliceX * i, (Picture1.ScaleHeight / 2) - (sliceY * val_motif)), vbRed, BF
            Else
                Picture1.Line (sliceX * oldi, (Picture1.ScaleHeight / 2))-(sliceX * i, (Picture1.ScaleHeight / 2) - (sliceY * val_motif)), vbBlue, BF
            End If
        End If
        
        
        'Picture1.Line (sliceX * oldi, (Picture1.ScaleHeight / 2) - (sliceY * OLD_val_motif))-(sliceX * i, (Picture1.ScaleHeight / 2) - (sliceY * val_motif)), vbRed
        'oldi = i
        'OLD_val_motif = val_motif
    
        oldi = i
    
    Next i
    '--------------------------------------------------------------------------
    
    'This is for an overlapping plot for the C+G signal
    '--------------------------------------------------------------------------
    'If KICP.Value = 1 Then
    '    oldi = 0
    '    For i = 1 To Len(ss) - cols
    '
    '        WindowIC = Mid(ss, i, Lungime_fereastra.Text)
    '        IC_motif = IC(WindowIC)
    '
    '        Picture1.Line (sliceX * oldi, Picture1.ScaleHeight - (sliceYIC * OLD_IC_motif))-(sliceX * i, Picture1.ScaleHeight - (sliceYIC * IC_motif)), vbBlack
    '        OLD_IC_motif = IC_motif
    '
    '        oldi = i
    '    Next i
    'End If
    '--------------------------------------------------------------------------
    
    'This is for an overlapping plot of self sequence alignment signal (IC)
    '--------------------------------------------------------------------------
    'If CGP.Value = 1 Then
    '    oldi = 0
    '    For i = 1 To Len(ss) - cols
    '
    '        WindowCG = Mid(ss, i, Lungime_fereastra.Text)
    '        CG_motif = CG(WindowCG)
    '
    '        Picture1.Line (sliceX * oldi, Picture1.ScaleHeight - (sliceYIC * OLD_CG_motif))-(sliceX * i, Picture1.ScaleHeight - (sliceYIC * CG_motif)), &H40C0&
    '        OLD_CG_motif = CG_motif
    '
    '        oldi = i
    '    '--------------------------------------------------------------------------
    '    Next i
    'End If
    
    'RealTime.Enabled = True
End Sub


Function IC(ByVal sequence As String) As Variant
    Dim count, i, u, total As Long
    Dim S1, s2 As String
    Dim max As Integer
    
    S1 = sequence
    max = Len(S1) - 1
        For u = 1 To max
            s2 = Mid(S1, u + 1)
            For i = 1 To Len(s2)
                If Mid(S1, i, 1) = Mid(s2, i, 1) Then
                    count = count + 1
                End If
            Next i
            total = total + (count / Len(s2) * 100)
            count = 0
        Next u
        
    IC = Round((total / max), 2)
End Function


Function CG(ByVal sequence As String) As Variant

    For i = 1 To Len(sequence)
    
        n = Mid(sequence, i, 1)
        If n = "A" Then a = a + 1
        If n = "C" Then c = c + 1
        If n = "G" Then g = g + 1
        If n = "T" Then t = t + 1

    Next i
    
    CG = (100 / (c + g + t + a)) * c + g
    
End Function


Private Sub secventata_Change()
    Label1.Caption = "Find motif in sequence [length: " & Len(secventata.Text) & "]:"
End Sub


Private Sub ShowVariants_Click(Index As Integer)

    ShowVariants(0).FontBold = False
    ShowVariants(1).FontBold = False
    ShowVariants(2).FontBold = False
    ShowVariants(3).FontBold = False
    ShowVariants(4).FontBold = False
    
    ShowVariants(Index).FontBold = True

    If Index = 0 Then
        MT_holder.Visible = True
        MV_holder.Visible = False
        MW_holder.Visible = False
        MP_holder.Visible = False
        IF_holder.Visible = False
    End If
    
    If Index = 1 Then
        MT_holder.Visible = False
        MV_holder.Visible = True
        MW_holder.Visible = False
        MP_holder.Visible = False
        IF_holder.Visible = False
    End If
    
    If Index = 2 Then
        MT_holder.Visible = False
        MV_holder.Visible = False
        MW_holder.Visible = True
        MP_holder.Visible = False
        IF_holder.Visible = False
    End If
    
    If Index = 3 Then
        MT_holder.Visible = False
        MV_holder.Visible = False
        MW_holder.Visible = False
        MP_holder.Visible = True
        IF_holder.Visible = False
    End If
    
    If Index = 4 Then
        MT_holder.Visible = False
        MV_holder.Visible = False
        MW_holder.Visible = False
        MP_holder.Visible = False
        IF_holder.Visible = True
    End If
    
End Sub


Private Sub Val_P_Click()
    Scann_Click
    'DoEvents
    'ShowVariants_Click (3)
End Sub


Private Sub Val_Z_Click()
    Scann_Click
    'DoEvents
    'ShowVariants_Click (2)
End Sub


Private Sub SLO_Click()
    Scann_Click
    DoEvents
    ShowVariants_Click (3)
End Sub


Private Sub ZSO_Click()
    Scann_Click
    DoEvents
    ShowVariants_Click (4)
End Sub


Private Sub WMatrix_Change()
    s = WMatrix.Text
    t = (Len(s) - Len(Replace(s, vbCrLf, ""))) / 2 '[vbcrlf represents two chr]
    Frame1.Caption = "The motif set (input): " & t & " motifs"
End Sub


Private Sub y_axis_Click()
    Scann_Click
End Sub

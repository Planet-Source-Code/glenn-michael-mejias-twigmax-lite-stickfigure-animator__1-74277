VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "TwigMax - Lite"
   ClientHeight    =   7005
   ClientLeft      =   165
   ClientTop       =   765
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   11640
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2940
      Left            =   1425
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   2910
      ScaleWidth      =   420
      TabIndex        =   11
      Top             =   330
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   5145
      Left            =   90
      ScaleHeight     =   5145
      ScaleWidth      =   1515
      TabIndex        =   3
      Top             =   90
      Width           =   1515
      Begin VB.CommandButton Command8 
         Caption         =   "Add Node"
         Height          =   480
         Left            =   0
         TabIndex        =   16
         Top             =   1245
         Width           =   1440
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   75
         Max             =   30
         Min             =   1
         TabIndex        =   14
         Top             =   3780
         Value           =   12
         Width           =   1365
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00000000&
         Height          =   330
         Left            =   90
         ScaleHeight     =   270
         ScaleWidth      =   1185
         TabIndex        =   12
         Top             =   300
         Width           =   1245
      End
      Begin VB.CommandButton Command3 
         Caption         =   ">>"
         Height          =   435
         Left            =   720
         TabIndex        =   9
         Top             =   1845
         Width           =   735
      End
      Begin VB.CommandButton Command5 
         Caption         =   "stop"
         Height          =   495
         Left            =   45
         TabIndex        =   7
         Top             =   4590
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "<<"
         Height          =   435
         Left            =   0
         TabIndex        =   10
         Top             =   1845
         Width           =   735
      End
      Begin VB.CommandButton Command4 
         Caption         =   "play"
         Height          =   495
         Left            =   45
         TabIndex        =   8
         Top             =   4110
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Next Frame (c)"
         Height          =   480
         Left            =   0
         TabIndex        =   6
         ToolTipText     =   "Shortcut 'C'"
         Top             =   2760
         Width           =   1440
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Add StickMan"
         Height          =   480
         Left            =   0
         TabIndex        =   5
         Top             =   780
         Width           =   1440
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Speed"
         Height          =   195
         Left            =   135
         TabIndex        =   15
         Top             =   3480
         Width           =   465
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Color"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   45
         Width           =   360
      End
      Begin VB.Label Label2 
         Caption         =   "CurrentFrame:0"
         Height          =   270
         Left            =   15
         TabIndex        =   4
         Top             =   2415
         Width           =   1410
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4095
      Top             =   7800
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   83
      Left            =   1635
      Top             =   7425
   End
   Begin VB.PictureBox picdraw 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   6735
      Left            =   1845
      ScaleHeight     =   445
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   641
      TabIndex        =   0
      Top             =   120
      Width           =   9675
   End
   Begin VB.CommandButton Command6 
      Enabled         =   0   'False
      Height          =   5355
      Left            =   15
      TabIndex        =   2
      Top             =   -30
      Width           =   1665
   End
   Begin VB.Label Label5 
      Caption         =   "Help: Right Click to add new segment"
      Height          =   405
      Left            =   135
      TabIndex        =   17
      Top             =   5535
      Width           =   1365
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   375
      Left            =   5880
      TabIndex        =   1
      Top             =   8880
      Width           =   975
   End
   Begin VB.Menu smfile 
      Caption         =   "File"
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu smabout 
      Caption         =   "?About"
      Begin VB.Menu smmikeytronixsoftwares 
         Caption         =   "Mikeytronix Softwares"
         Begin VB.Menu smglennmichaelmejias 
            Caption         =   "Glenn Michael Mejias"
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type segtype
    segcount As Long
    angle As Double
    scale As Double
    ind As Long
    x As Double
    y As Double
    con As Long
    w As Double
    col As Long
    typ As String
End Type
Private Type indanim
    segcount As New Collection
    angle As New Collection
    scale As New Collection
    ind As New Collection
    x As New Collection
    y As New Collection
    con As New Collection
    w As New Collection
    col As New Collection
    typ As New Collection
    figind As New Collection
End Type
Private Type translate_prop
    ind As Long
    tangle As Double
    tlength As Double
    tx As Double
    ty As Double
End Type
Dim translate_arr(1000) As translate_prop
Dim arr(200, 1000) As segtype
Dim eframes(2000) As indanim
Dim currentindex As Long
Dim currentelement As Long
Dim currentframe As Long
Dim maxframe As Long
Dim elementcount As Long
Dim tempcount As Long
Private Sub Command1_Click()
    Dim a As Long
    Dim b As Long
    For b = 0 To elementcount
        For a = 0 To get_segcount(b, 0)
            eframes(currentframe).x.Add get_x(b, a)
            eframes(currentframe).y.Add get_y(b, a)
            eframes(currentframe).ind.Add a
            eframes(currentframe).figind.Add b
            eframes(currentframe).segcount.Add get_segcount(b, 0)
            eframes(currentframe).con.Add get_con(b, a)
            eframes(currentframe).typ.Add get_typ(b, a)
        Next a
    Next b
    currentframe = currentframe + 1
    If maxframe < currentframe Then
        maxframe = currentframe
    End If
    Label2.Caption = "CurrentFrame: " & currentframe
    animate_element currentframe
End Sub
Public Sub animate_element(frme As Long)
    Dim a As Long
    Dim v As Variant
    Dim x As Double
    Dim y As Double
    Dim w As Double
    Dim figind As Long
    For Each v In eframes(frme).ind
        a = a + 1
        x = eframes(frme).x(a)
        y = eframes(frme).y(a)
        figind = eframes(frme).figind(a)
        set_x x, figind, CLng(v)
        set_y y, figind, CLng(v)
        set_con eframes(frme).con(a), figind, CLng(v)
        'set_ind a - 1, figind, CLng(v)
        set_segcount CLng(eframes(frme).segcount(a)), figind, CLng(v)
        set_typ eframes(frme).typ(a), figind, CLng(v)
    Next
    draw_element picdraw
End Sub
Private Sub Command2_Click()
    If currentframe - 1 < 0 Then
        Exit Sub
    End If
    currentframe = currentframe - 1
    Label2.Caption = "CurrentFrame: " & currentframe
    animate_element currentframe
End Sub

Private Sub Command3_Click()
    'maxframe = maxframe + 1
    If currentframe + 1 > maxframe Then
        Exit Sub
    End If
    currentframe = currentframe + 1
    Label2.Caption = "CurrentFrame: " & currentframe
    animate_element currentframe
End Sub

Private Sub Command4_Click()
    currentframe = 0
    Timer1.Enabled = True
    Timer2.Enabled = True
End Sub
Private Sub Command5_Click()
    Timer1.Enabled = False
    Timer2.Enabled = False
    currentframe = 0
    animate_element currentframe
    Label1.Caption = 0
    Label2.Caption = "CurrentFrame:" & 0
End Sub
Public Sub add_stickman()
    Dim a As Long
    a = elementcount
    arr(a, 0).x = 436
    arr(a, 0).y = 282
    arr(a, 0).con = 0
    arr(a, 0).w = 23
    arr(a, 0).col = 0
    arr(a, 0).typ = "l"
    arr(a, 0).ind = 0
    
    arr(a, 1).x = 436
    arr(a, 1).y = 238
    arr(a, 1).con = 0
    arr(a, 1).w = 23
    arr(a, 1).col = 0
    arr(a, 1).typ = "l"
    arr(a, 1).ind = 0
    
    arr(a, 2).x = 436
    arr(a, 2).y = 195
    arr(a, 2).con = 1
    arr(a, 2).w = 23
    arr(a, 2).col = 0
    arr(a, 2).typ = "l"
    arr(a, 2).ind = 0
    
    arr(a, 3).x = 414
    arr(a, 3).y = 247
    arr(a, 3).con = 2
    arr(a, 3).w = 23
    arr(a, 3).col = 0
    arr(a, 3).typ = "l"
    arr(a, 3).ind = 0
    
    arr(a, 4).x = 459
    arr(a, 4).y = 246
    arr(a, 4).con = 2
    arr(a, 4).w = 23
    arr(a, 4).col = 0
    arr(a, 4).typ = "l"
    arr(a, 4).ind = 0
    
    arr(a, 5).x = 404
    arr(a, 5).y = 292
    arr(a, 5).con = 3
    arr(a, 5).w = 23
    arr(a, 5).col = 0
    arr(a, 5).typ = "l"
    arr(a, 5).ind = 0
    
    arr(a, 6).x = 470
    arr(a, 6).y = 291
    arr(a, 6).con = 4
    arr(a, 6).w = 23
    arr(a, 6).col = 0
    arr(a, 6).typ = "l"
    arr(a, 6).ind = 0
    
    arr(a, 7).x = 420
    arr(a, 7).y = 359
    arr(a, 7).con = 0
    arr(a, 7).w = 27
    arr(a, 7).col = 0
    arr(a, 7).typ = "l"
    arr(a, 7).ind = 0
    
    arr(a, 8).x = 418
    arr(a, 8).y = 424
    arr(a, 8).con = 7
    arr(a, 8).w = 27
    arr(a, 8).col = 0
    arr(a, 8).typ = "l"
    arr(a, 8).ind = 0
    
    arr(a, 9).x = 454
    arr(a, 9).y = 358
    arr(a, 9).con = 0
    arr(a, 9).w = 27
    arr(a, 9).col = 0
    arr(a, 9).typ = "l"
    arr(a, 9).ind = 0
    
    arr(a, 10).x = 456
    arr(a, 10).y = 423
    arr(a, 10).con = 9
    arr(a, 10).w = 27
    arr(a, 10).col = 0
    arr(a, 10).typ = "l"
    arr(a, 10).ind = 0
    
    arr(a, 11).x = 412
    arr(a, 11).y = 430
    arr(a, 11).con = 8
    arr(a, 11).w = 27
    arr(a, 11).col = 0
    arr(a, 11).typ = "l"
    arr(a, 11).ind = 0
    
    arr(a, 12).x = 464
    arr(a, 12).y = 428
    arr(a, 12).con = 10
    arr(a, 12).w = 27
    arr(a, 12).col = 0
    arr(a, 12).typ = "l"
    arr(a, 12).ind = 0
    
    arr(a, 13).x = 436
    arr(a, 13).y = 172
    arr(a, 13).con = 2
    arr(a, 13).w = 19
    arr(a, 13).col = 0
    arr(a, 13).typ = "l"
    arr(a, 13).ind = 0
    
    arr(a, 14).x = 436
    arr(a, 14).y = 149
    arr(a, 14).con = 13
    arr(a, 14).w = 34
    arr(a, 14).col = 0
    arr(a, 14).typ = "c"
    arr(a, 14).ind = 0
    
    arr(a, 0).segcount = 14
    
    elementcount = elementcount + 1
End Sub

Private Sub Command7_Click()
    add_stickman
    draw_element picdraw
End Sub

Private Sub Command8_Click()
    a = elementcount
    arr(a, 0).x = 436
    arr(a, 0).y = 282
    arr(a, 0).con = 0
    arr(a, 0).w = 23
    arr(a, 0).col = 0
    arr(a, 0).typ = "l"
    arr(a, 0).ind = 0
    
    arr(a, 0).segcount = 0
    
    elementcount = elementcount + 1
    draw_element picdraw
End Sub

Private Sub exit_Click()
    End
End Sub

Private Sub Form_Load()
    add_stickman
    draw_element picdraw
End Sub
Private Function draw_circle(picbox As PictureBox, x1 As Double, y1 As Double, x2 As Double, y2 As Double, w As Double)
    Dim d As Double
    picbox.DrawWidth = w
    d = Sqr(((x2 - x1) * (x2 - x1)) + ((y2 - y1) * (y2 - y1)))
    picbox.Circle (x2 / 2 + x1 / 2, y2 / 2 + y1 / 2), d / 2
End Function
Public Sub addsegment()
    Dim a As Long
    Dim x As Double
    Dim y As Double
    Dim w As Double
    Dim col As Long
    Dim typ As String
    x = get_x(currentelement, currentindex)
    y = get_y(currentelement, currentindex)
    w = get_w(currentelement, currentindex)
    col = get_col(currentelement, currentindex)
    typ = get_typ(currentelement, currentindex)
    set_segcount get_segcount(currentelement, 0) + 1, currentelement, 0
    set_x x, currentelement, get_segcount(currentelement, 0)
    set_y y, currentelement, get_segcount(currentelement, 0)
    set_w w, currentelement, get_segcount(currentelement, 0)
    set_typ typ, currentelement, get_segcount(currentelement, 0)
    set_con currentindex, currentelement, get_segcount(currentelement, 0)
    currentindex = get_segcount(currentelement, 0)
    draw_element picdraw
End Sub
Public Sub draw_element(picbox As PictureBox)
    Dim a As Long
    Dim b As Long
    Dim x1 As Double
    Dim x2 As Double
    Dim y1 As Double
    Dim y2 As Double
    Dim x As Double
    Dim y As Double
    Dim tx As Double
    Dim ty As Double
    Dim w As Double
    Dim col As Long
    Dim con As Long
    Dim typ As String
    picbox.Refresh
    picbox.Cls
    For b = 0 To elementcount - 1
        For a = 0 To get_segcount(b, 0)
            x1 = get_x(b, a)
            y1 = get_y(b, a)
            con = get_con(b, a)
            w = get_w(b, a)
            col = get_col(b, a)
            typ = get_typ(b, a)
            x2 = get_x(b, con)
            y2 = get_y(b, con)
            If typ = "l" Then
                draw_line picbox, x1, y1, x2, y2, col, w
            ElseIf typ = "c" Then
                draw_circle picbox, x1, y1, x2, y2, w
            End If
        Next a
        For a = 0 To get_segcount(b, 0)
            x1 = get_x(b, a)
            y1 = get_y(b, a)
            picbox.FillStyle = 0
            picbox.DrawWidth = 1
            If a > 0 Then
                picbox.ForeColor = vbRed
                picbox.FillColor = vbRed
                picbox.Circle (x1, y1), 4
            Else
                picbox.ForeColor = RGB(255, 200, 0)
                picbox.FillColor = RGB(255, 200, 0)
                picbox.Circle (x1, y1), 5
            End If
        Next a
    Next b
End Sub
Public Sub draw_line(picbox As PictureBox, x1 As Double, y1 As Double, x2 As Double, y2 As Double, col As Long, w As Double)
    picbox.ForeColor = col
    picbox.DrawWidth = w
    picbox.Line (x1, y1)-(x2, y2)
End Sub
Public Function get_x(eind As Long, sind As Long) As Double
    get_x = arr(eind, sind).x
End Function
Public Function get_y(eind As Long, sind As Long) As Double
    get_y = arr(eind, sind).y
End Function
Public Function get_w(eind As Long, sind As Long) As Double
    get_w = arr(eind, sind).w
End Function
Public Function get_con(eind As Long, sind As Long) As Long
    get_con = arr(eind, sind).con
End Function
Public Function get_col(eind As Long, sind As Long) As Long
    get_col = arr(eind, sind).col
End Function
Public Function get_typ(eind As Long, sind As Long) As String
    get_typ = arr(eind, sind).typ
End Function
Public Function get_ind(eind As Long, sind As Long) As String
    get_ind = arr(eind, sind).ind
End Function
Public Function get_segcount(eind As Long, sind As Long) As Long
    get_segcount = arr(eind, sind).segcount
End Function

Public Sub set_x(itm As Double, eind As Long, sind As Long)
    arr(eind, sind).x = itm
End Sub
Public Sub set_y(itm As Double, eind As Long, sind As Long)
    arr(eind, sind).y = itm
End Sub
Public Sub set_w(itm As Double, eind As Long, sind As Long)
    arr(eind, sind).w = itm
End Sub
Public Sub set_con(itm As Long, eind As Long, sind As Long)
    arr(eind, sind).con = itm
End Sub
Public Sub set_col(itm As Long, eind As Long, sind As Long)
    arr(eind, sind).col = itm
End Sub
Public Sub set_typ(itm As String, eind As Long, sind As Long)
    arr(eind, sind).typ = itm
End Sub
Public Sub set_ind(itm As String, eind As Long, sind As Long)
    arr(eind, sind).ind = itm
End Sub
Public Sub set_segcount(itm As Long, eind As Long, sind As Long)
    arr(eind, 0).segcount = itm
End Sub

Private Sub HScroll1_Scroll()
    Timer1.Interval = 1000 / Val(HScroll1.Value)
End Sub

Private Sub picdraw_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 67 Then
        Command1_Click
    End If
End Sub

Private Sub picdraw_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim a As Long
    Dim b As Long
    Dim y1 As Double
    Dim y2 As Double
    Const mwidth = 6
    currentindex = -1
    For b = elementcount To 0 Step -1
        For a = get_segcount(b, 0) To 0 Step -1
            x2 = get_x(b, a)
            y2 = get_y(b, a)
            If (x > x2 - mwidth And x < x2 + mwidth) And (y > y2 - mwidth And y < y2 + mwidth) Then
                currentindex = a
                currentelement = b
                GoTo jump_here
            End If
        Next a
    Next b
    Exit Sub
jump_here:
    get_indexs
    storelocation CDbl(x), CDbl(y)
    If Button = 2 Then
        addsegment
    End If
End Sub
Public Sub get_indexs()
    Dim b As Long
    Dim a As Long
    Dim tmp As Long
    tempcount = 0
    translate_arr(tempcount).ind = currentindex
    tempcount = tempcount + 1
    b = 0
    Do While b < tempcount
        If b > get_segcount(currentelement, 0) Then GoTo jump_here
        For a = 1 To get_segcount(currentelement, 0)
            tmp = get_con(currentelement, a)
            If translate_arr(b).ind = tmp Then
                translate_arr(tempcount).ind = a
                tempcount = tempcount + 1
            End If
        Next a
        b = b + 1
    Loop
jump_here:
    get_angle
End Sub
Public Sub storelocation(x As Double, y As Double)
    Dim a As Long
    Dim x1 As Double
    Dim y1 As Double
    For a = 0 To tempcount - 1
        x1 = get_x(currentelement, translate_arr(a).ind)
        y1 = get_y(currentelement, translate_arr(a).ind)
        translate_arr(a).tx = x1 - x
        translate_arr(a).ty = y1 - y
    Next a
End Sub
Public Sub movefigure(x As Double, y As Double)
    Dim a As Long
    Dim b As Long
    For a = 0 To tempcount - 1
        set_x x + translate_arr(a).tx, currentelement, translate_arr(a).ind
        set_y y + translate_arr(a).ty, currentelement, translate_arr(a).ind
    Next a
End Sub
Public Sub get_angle()
    Dim a As Long
    Dim angle As Double
    Dim length As Double
    Dim x1 As Double
    Dim y1 As Double
    Dim x2 As Double
    Dim y2 As Double
    Dim rootangle As Double
    Dim con As Long
    con = get_con(currentelement, currentindex)
    x1 = get_x(currentelement, con)
    y1 = get_y(currentelement, con)
    x2 = get_x(currentelement, currentindex)
    y2 = get_y(currentelement, currentindex)
    rootangle = getangle(x1, y1, x2, y2)
    For a = 0 To tempcount - 1
        x2 = get_x(currentelement, translate_arr(a).ind)
        y2 = get_y(currentelement, translate_arr(a).ind)
        angle = getangle(x1, y1, x2, y2)
        length = getlength(x1, y1, x2, y2)
        translate_arr(a).tangle = angle - rootangle
        translate_arr(a).tlength = length
    Next a
End Sub
Private Sub picdraw_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If currentindex = -1 Then
        Exit Sub
    End If
    If Button = 1 Then
        If currentindex = 0 Then
            movefigure CDbl(x), CDbl(y)
        Else
            rotate_segment CDbl(x), CDbl(y)
        End If
        draw_element picdraw
    ElseIf Button = 2 Then
        set_x Val(x), currentelement, currentindex
        set_y Val(y), currentelement, currentindex
        draw_element picdraw
    End If
End Sub

Public Sub rotate_segment(x As Double, y As Double)
    Dim x1 As Double
    Dim y1 As Double
    Dim x2 As Double
    Dim y2 As Double
    Dim con As Long
    Dim length As Double
    Dim angle As Double
    Dim a As Long
    Dim b As Long
    Dim p1 As POINTAPI
    con = get_con(currentelement, currentindex)
    x2 = get_x(currentelement, con)
    y2 = get_y(currentelement, con)
    For a = 0 To tempcount - 1
        length = translate_arr(a).tlength
        angle = getangle(x2, y2, x, y) + translate_arr(a).tangle
        p1 = rotatepoint(angle, length, x2, y2)
        set_x p1.x, currentelement, translate_arr(a).ind
        set_y p1.y, currentelement, translate_arr(a).ind
    Next a
End Sub
Public Function getlength(x1 As Double, y1 As Double, x2 As Double, y2 As Double) As Double
    getlength = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)
End Function

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Dim temp As Long
        temp = Picture2.Point(x, y)
        If temp < 0 Then
            temp = 0
        End If
        Picture3.BackColor = temp
        
        Dim a As Long
        For a = 0 To get_segcount(currentelement, 0)
            set_col Picture3.BackColor, currentelement, a
        Next a
        draw_element picdraw
    End If
End Sub

Private Sub Picture3_Click()
    If Picture2.Visible = True Then
        Picture2.Visible = False
    ElseIf Picture2.Visible = False Then
        Picture2.Visible = True
    End If
End Sub

Private Sub Timer1_Timer()
    If currentframe > maxframe Then
        currentframe = 0
    End If
    currentframe = currentframe + 1
    Label2.Caption = "CurrentFrame: " & currentframe
    animate_element currentframe
End Sub
Private Sub Timer2_Timer()
    Label1.Caption = Val(Label1.Caption + 1)
End Sub

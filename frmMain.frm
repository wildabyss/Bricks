VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bricks Classic"
   ClientHeight    =   8655
   ClientLeft      =   4020
   ClientTop       =   1485
   ClientWidth     =   6735
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   6735
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   1980
      Left            =   4560
      TabIndex        =   5
      Top             =   960
      Width           =   1980
      Begin VB.Shape shpPreview 
         BorderColor     =   &H000000C0&
         FillColor       =   &H00FF8080&
         FillStyle       =   0  'Solid
         Height          =   405
         Index           =   3
         Left            =   1335
         Top             =   840
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Shape shpPreview 
         BorderColor     =   &H000000C0&
         FillColor       =   &H00FF8080&
         FillStyle       =   0  'Solid
         Height          =   405
         Index           =   2
         Left            =   930
         Top             =   840
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Shape shpPreview 
         BorderColor     =   &H000000C0&
         FillColor       =   &H00FF8080&
         FillStyle       =   0  'Solid
         Height          =   405
         Index           =   1
         Left            =   525
         Top             =   840
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Shape shpPreview 
         BorderColor     =   &H000000C0&
         FillColor       =   &H00FF8080&
         FillStyle       =   0  'Solid
         Height          =   405
         Index           =   0
         Left            =   120
         Top             =   840
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Shape shpBrick 
         BorderColor     =   &H000000C0&
         FillColor       =   &H00FF8080&
         FillStyle       =   0  'Solid
         Height          =   405
         Index           =   1
         Left            =   0
         Top             =   7695
         Width           =   405
      End
   End
   Begin VB.ListBox lstSpeed 
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      ItemData        =   "frmMain.frx":0152
      Left            =   4560
      List            =   "frmMain.frx":0165
      TabIndex        =   3
      Top             =   6000
      Width           =   1935
   End
   Begin VB.Timer timBricks 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   5280
      Top             =   8040
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   8100
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4050
      Begin VB.Shape shpBrick 
         BorderColor     =   &H000000C0&
         FillColor       =   &H00FF8080&
         FillStyle       =   0  'Solid
         Height          =   405
         Index           =   0
         Left            =   0
         Top             =   7695
         Visible         =   0   'False
         Width           =   405
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Next Piece:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Speed"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Score:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Menu mnuStart 
      Caption         =   "Start (F5)"
   End
   Begin VB.Menu mnuPause 
      Caption         =   "Pause (Esc)"
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim booIsMoving As Boolean
Dim booGrid(9, 19) As Boolean
Dim shpGrid(9, 19) As Shape
Dim lngCounter As Long  'shpBrick's index
Dim lngScore As Long
Dim lngSpeed As Long
Dim lngMax As Long  'max y
Dim shpLeft As Shape, shpRight As Shape, shpBottom As Shape 'left=leftiest, right=rightiest, bottom=lowest, blocks
Dim lngBrick1 As Long, lngBrick2 As Long  'shape1=the moving brick, shape2=brick in preview
Dim lngType As Long     'shape of the brick

'determine whether the brick can rotate
Private Function detObstruction(a As Long, b As Long) As Boolean
    Dim x, y As Long
    x = a / 405: y = b / 405
    On Error Resume Next        'IF YOU'VE TIME, RESOLVE THIS IMMEDIATELY!!
    If booGrid(x - 1, y - 1) Then
        detObstruction = True
    ElseIf booGrid(x, y - 1) Then
        detObstruction = True
    ElseIf booGrid(x + 1, y - 1) Then
        detObstruction = True
    ElseIf booGrid(x - 1, y) Then
        detObstruction = True
    ElseIf booGrid(x + 1, y) Then
        detObstruction = True
    ElseIf booGrid(x - 1, y + 1) Then
        detObstruction = True
    ElseIf booGrid(x, y + 1) Then
        detObstruction = True
    ElseIf booGrid(x + 1, y + 1) Then
        detObstruction = True
    Else
        detObstruction = False
    End If
End Function

'determine if the brick is touching another brick in any way
Private Function detProximity(kind As String) As Boolean
    Dim b As Long
    Dim answer As Boolean
    Select Case kind
        Case "right"
            For b = (lngCounter - 3) To lngCounter
                If booGrid((shpBrick(b).Left + 405) / 405, shpBrick(b).Top / 405) Then answer = True
            Next
        Case "left"
            For b = (lngCounter - 3) To lngCounter
                If booGrid((shpBrick(b).Left - 405) / 405, shpBrick(b).Top / 405) Then answer = True
            Next
        Case "bottom"
            For b = (lngCounter - 3) To lngCounter
                If booGrid(shpBrick(b).Left / 405, (shpBrick(b).Top + 405) / 405) Then answer = True
            Next
    End Select
    If answer Then
        detProximity = True
    Else
        detProximity = False
    End If
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim b As Long
    
    If KeyCode = vbKeyDown Then
        '1. check IsMoving
        If booIsMoving = False Then Exit Sub
        '2. check if it's at the bottom
        If shpBottom.Top = 7695 Then Exit Sub
        '3. check if it touches any other blocks
        'For b = (lngCounter - 3) To lngCounter
        '    If booGrid(shpBrick(b).Left / 405, (shpBrick(b).Top + 405) / 405) Then Exit Sub
        'Next
        If detProximity("bottom") Then Exit Sub
        '4. if not, move down 405
        shpBrick(lngCounter).Top = shpBrick(lngCounter).Top + 405
        shpBrick(lngCounter - 1).Top = shpBrick(lngCounter - 1).Top + 405
        shpBrick(lngCounter - 2).Top = shpBrick(lngCounter - 2).Top + 405
        shpBrick(lngCounter - 3).Top = shpBrick(lngCounter - 3).Top + 405
        
    ElseIf KeyCode = vbKeyRight Then
        '0. check IsMoving
        If booIsMoving = False Then Exit Sub
        '1. check if it's at the right brink
        If shpRight.Left = 3645 Then Exit Sub
        '2. check if it touches any other blocks
        If detProximity("right") Then Exit Sub
        '3. if not, move right 405
        shpBrick(lngCounter).Left = shpBrick(lngCounter).Left + 405
        shpBrick(lngCounter - 1).Left = shpBrick(lngCounter - 1).Left + 405
        shpBrick(lngCounter - 2).Left = shpBrick(lngCounter - 2).Left + 405
        shpBrick(lngCounter - 3).Left = shpBrick(lngCounter - 3).Left + 405
        
    ElseIf KeyCode = vbKeyLeft Then
        '0. check IsMoving
        If booIsMoving = False Then Exit Sub
        '1. check if it's at the left brink
        If shpLeft.Left = 0 Then Exit Sub
        '2. check if it touches any other blocks
        If detProximity("left") Then Exit Sub
        '3. if not, move left 405
        shpBrick(lngCounter).Left = shpBrick(lngCounter).Left - 405
        shpBrick(lngCounter - 1).Left = shpBrick(lngCounter - 1).Left - 405
        shpBrick(lngCounter - 2).Left = shpBrick(lngCounter - 2).Left - 405
        shpBrick(lngCounter - 3).Left = shpBrick(lngCounter - 3).Left - 405
        
    ElseIf KeyCode = vbKeySpace Then    'change shape
        'transformation
        Select Case lngBrick1
            Case 0
                Select Case lngType
                    Case 0  'upright
                        If (shpRight.Left = 3645) And (Not (detObstruction(shpBrick(lngCounter - 2).Left - 405, shpBrick(lngCounter - 2).Top))) Then
                            
                        ElseIf detProximity("right") And (Not (detObstruction(shpBrick(lngCounter - 2).Left - 405, shpBrick(lngCounter - 2).Top))) Then
                            
                        ElseIf detObstruction(shpBrick(lngCounter - 2).Left, shpBrick(lngCounter - 2).Top) Then
                            Exit Sub
                        Else
                            shpBrick(lngCounter - 1).Top = shpBrick(lngCounter - 3).Top
                            shpBrick(lngCounter).Left = shpBrick(lngCounter - 1).Left
                            shpBrick(lngCounter).Top = shpBrick(lngCounter - 2).Top
                            shpBrick(lngCounter - 3).Top = shpBrick(lngCounter - 2).Top
                            shpBrick(lngCounter - 3).Left = shpBrick(lngCounter - 3).Left + 405
                            Set shpRight = shpBrick(lngCounter - 3)
                            Set shpBottom = shpBrick(lngCounter)
                            lngType = 1
                        End If
                    Case 1  'belly-up
                        If shpBottom.Top = 7695 Then
                            Exit Sub
                        ElseIf detObstruction(shpBrick(lngCounter - 2).Left, shpBrick(lngCounter - 2).Top) Then
                            Exit Sub
                        Else
                            shpBrick(lngCounter - 1).Left = shpBrick(lngCounter - 3).Left
                            shpBrick(lngCounter).Left = shpBrick(lngCounter - 2).Left
                            shpBrick(lngCounter).Top = shpBrick(lngCounter - 1).Top
                            shpBrick(lngCounter - 3).Left = shpBrick(lngCounter - 2).Left
                            shpBrick(lngCounter - 3).Top = shpBrick(lngCounter - 3).Top + 405
                            Set shpRight = shpBrick(lngCounter - 1)
                            Set shpBottom = shpBrick(lngCounter - 3)
                            Set shpLeft = shpBrick(lngCounter - 3)
                            lngType = 2
                        End If
                    Case 2  'upside-down and upright
                        If shpLeft.Left = 0 Then
                            
                        ElseIf detProximity("left") And (Not (detObstruction(shpBrick(lngCounter - 2).Left + 405, shpBrick(lngCounter - 2).Top))) Then
                        
                        ElseIf detObstruction(shpBrick(lngCounter - 2).Left, shpBrick(lngCounter - 2).Top) Then
                            Exit Sub
                        Else
                            shpBrick(lngCounter - 1).Top = shpBrick(lngCounter - 3).Top
                            shpBrick(lngCounter).Left = shpBrick(lngCounter - 1).Left
                            shpBrick(lngCounter).Top = shpBrick(lngCounter - 2).Top
                            shpBrick(lngCounter - 3).Left = shpBrick(lngCounter - 2).Left - 405
                            shpBrick(lngCounter - 3).Top = shpBrick(lngCounter - 2).Top
                            Set shpRight = shpBrick(lngCounter - 1)
                            Set shpBottom = shpBrick(lngCounter - 1)
                            Set shpLeft = shpBrick(lngCounter - 3)
                            lngType = 3
                        End If
                    Case 3  'upside-down and belly-up
                        If detObstruction(shpBrick(lngCounter - 2).Left, shpBrick(lngCounter - 2).Top) Then
                            Exit Sub
                        Else
                            shpBrick(lngCounter - 1).Left = shpBrick(lngCounter - 3).Left
                            shpBrick(lngCounter).Left = shpBrick(lngCounter - 2).Left
                            shpBrick(lngCounter).Top = shpBrick(lngCounter - 1).Top
                            shpBrick(lngCounter - 3).Left = shpBrick(lngCounter - 2).Left
                            shpBrick(lngCounter - 3).Top = shpBrick(lngCounter - 2).Top - 405
                            Set shpRight = shpBrick(lngCounter - 3)
                            Set shpBottom = shpBrick(lngCounter - 1)
                            Set shpLeft = shpBrick(lngCounter - 1)
                            lngType = 0
                        End If
                End Select
            Case 1  'Z
                Select Case lngType
                    Case 0
                        If shpLeft.Left = 0 Then
                            
                        ElseIf detProximity("left") And (Not (detObstruction(shpBrick(lngCounter - 2).Left + 405, shpBrick(lngCounter - 2).Top))) Then
                            
                        ElseIf detObstruction(shpBrick(lngCounter - 2).Left, shpBrick(lngCounter - 2).Top) Then
                            Exit Sub
                        Else
                            shpBrick(lngCounter - 3).Top = shpBrick(lngCounter - 2).Top + 405
                            shpBrick(lngCounter - 1).Left = shpBrick(lngCounter - 2).Left
                            shpBrick(lngCounter - 1).Top = shpBrick(lngCounter - 3).Top
                            shpBrick(lngCounter).Left = shpBrick(lngCounter - 2).Left - 405
                            shpBrick(lngCounter).Top = shpBrick(lngCounter - 2).Top
                            Set shpLeft = shpBrick(lngCounter)
                            Set shpRight = shpBrick(lngCounter - 3)
                            Set shpBottom = shpBrick(lngCounter - 1)
                            lngType = 1
                        End If
                    Case 1
                        If detObstruction(shpBrick(lngCounter - 2).Left, shpBrick(lngCounter - 2).Top) Then
                            Exit Sub
                        Else
                            shpBrick(lngCounter - 3).Top = shpBrick(lngCounter - 2).Top - 405
                            shpBrick(lngCounter - 1).Left = shpBrick(lngCounter - 3).Left
                            shpBrick(lngCounter - 1).Top = shpBrick(lngCounter - 2).Top
                            shpBrick(lngCounter).Left = shpBrick(lngCounter - 2).Left
                            shpBrick(lngCounter).Top = shpBrick(lngCounter - 2).Top + 405
                            Set shpLeft = shpBrick(lngCounter - 2)
                            Set shpRight = shpBrick(lngCounter - 1)
                            Set shpBottom = shpBrick(lngCounter)
                            lngType = 0
                        End If
                End Select
            Case 2
                Select Case lngType
                    Case 0
                        If shpLeft.Left = 0 Then
                            
                        ElseIf detProximity("left") And (Not (detObstruction(shpBrick(lngCounter - 2).Left + 405, shpBrick(lngCounter - 2).Top))) Then
                        
                        ElseIf detObstruction(shpBrick(lngCounter - 2).Left, shpBrick(lngCounter - 2).Top) Then
                            Exit Sub
                        Else
                            shpBrick(lngCounter - 3).Top = shpBrick(lngCounter - 2).Top
                            shpBrick(lngCounter - 3).Left = shpBrick(lngCounter).Left
                            shpBrick(lngCounter - 1).Left = shpBrick(lngCounter - 2).Left - 405
                            shpBrick(lngCounter - 1).Top = shpBrick(lngCounter - 2).Top
                            shpBrick(lngCounter).Left = shpBrick(lngCounter - 1).Left
                            Set shpLeft = shpBrick(lngCounter - 1)
                            Set shpRight = shpBrick(lngCounter - 3)
                            Set shpBottom = shpBrick(lngCounter)
                            lngType = 1
                        End If
                    Case 1
                        If detObstruction(shpBrick(lngCounter - 2).Left, shpBrick(lngCounter - 2).Top) Then
                            Exit Sub
                        Else
                            shpBrick(lngCounter - 3).Top = shpBrick(lngCounter).Top
                            shpBrick(lngCounter - 3).Left = shpBrick(lngCounter - 2).Left
                            shpBrick(lngCounter - 1).Left = shpBrick(lngCounter - 2).Left
                            shpBrick(lngCounter - 1).Top = shpBrick(lngCounter - 2).Top - 405
                            shpBrick(lngCounter).Top = shpBrick(lngCounter - 1).Top
                            Set shpLeft = shpBrick(lngCounter)
                            Set shpRight = shpBrick(lngCounter - 1)
                            Set shpBottom = shpBrick(lngCounter - 3)
                            lngType = 2
                        End If
                    Case 2
                        If (shpRight.Left = 3645) And (Not (detObstruction(shpBrick(lngCounter - 2).Left - 405, shpBrick(lngCounter - 2).Top))) Then
                            
                        ElseIf detProximity("right") And (Not (detObstruction(shpBrick(lngCounter - 2).Left - 405, shpBrick(lngCounter - 2).Top))) Then
                            
                        ElseIf detObstruction(shpBrick(lngCounter - 2).Left, shpBrick(lngCounter - 2).Top) Then
                            Exit Sub
                        Else
                            shpBrick(lngCounter - 3).Top = shpBrick(lngCounter - 2).Top
                            shpBrick(lngCounter - 3).Left = shpBrick(lngCounter).Left
                            shpBrick(lngCounter - 1).Left = shpBrick(lngCounter - 2).Left + 405
                            shpBrick(lngCounter - 1).Top = shpBrick(lngCounter - 2).Top
                            shpBrick(lngCounter).Left = shpBrick(lngCounter - 1).Left
                            Set shpLeft = shpBrick(lngCounter - 3)
                            Set shpRight = shpBrick(lngCounter - 1)
                            Set shpBottom = shpBrick(lngCounter - 3)
                            lngType = 3
                        End If
                    Case 3
                        If shpBottom.Top = 7695 Then
                            Exit Sub
                        ElseIf detObstruction(shpBrick(lngCounter - 2).Left, shpBrick(lngCounter - 2).Top) Then
                            Exit Sub
                        Else
                            shpBrick(lngCounter - 3).Top = shpBrick(lngCounter).Top
                            shpBrick(lngCounter - 3).Left = shpBrick(lngCounter - 2).Left
                            shpBrick(lngCounter - 1).Left = shpBrick(lngCounter - 2).Left
                            shpBrick(lngCounter - 1).Top = shpBrick(lngCounter - 2).Top + 405
                            shpBrick(lngCounter).Top = shpBrick(lngCounter - 1).Top
                            Set shpLeft = shpBrick(lngCounter - 1)
                            Set shpRight = shpBrick(lngCounter)
                            Set shpBottom = shpBrick(lngCounter - 1)
                            lngType = 0
                        End If
                End Select
            Case 3  'inverse Z
                Select Case lngType
                    Case 0
                        If shpLeft.Left = 0 Then
                            
                        ElseIf detProximity("left") And (Not (detObstruction(shpBrick(lngCounter - 2).Left + 405, shpBrick(lngCounter - 2).Top))) Then
                            
                        ElseIf detObstruction(shpBrick(lngCounter - 2).Left, shpBrick(lngCounter - 2).Top) Then
                            Exit Sub
                        Else
                            shpBrick(lngCounter - 3).Top = shpBrick(lngCounter - 2).Top
                            shpBrick(lngCounter - 3).Left = shpBrick(lngCounter - 1).Left
                            shpBrick(lngCounter - 1).Left = shpBrick(lngCounter - 2).Left
                            shpBrick(lngCounter - 1).Top = shpBrick(lngCounter).Top
                            shpBrick(lngCounter).Left = shpBrick(lngCounter - 2).Left - 405
                            Set shpLeft = shpBrick(lngCounter)
                            Set shpRight = shpBrick(lngCounter - 3)
                            Set shpBottom = shpBrick(lngCounter)
                            lngType = 1
                        End If
                    Case 1
                        If detObstruction(shpBrick(lngCounter - 2).Left, shpBrick(lngCounter - 2).Top) Then
                            Exit Sub
                        Else
                            shpBrick(lngCounter - 3).Top = shpBrick(lngCounter - 2).Top - 405
                            shpBrick(lngCounter - 3).Left = shpBrick(lngCounter - 2).Left
                            shpBrick(lngCounter - 1).Left = shpBrick(lngCounter - 2).Left + 405
                            shpBrick(lngCounter - 1).Top = shpBrick(lngCounter - 2).Top
                            shpBrick(lngCounter).Left = shpBrick(lngCounter - 1).Left
                            Set shpLeft = shpBrick(lngCounter - 2)
                            Set shpRight = shpBrick(lngCounter)
                            Set shpBottom = shpBrick(lngCounter)
                            lngType = 0
                        End If
                End Select
            Case 5
                Select Case lngType
                    Case 0
                        If shpBottom.Top = 7695 Then
                            Exit Sub
                        ElseIf detObstruction(shpBrick(lngCounter - 1).Left, shpBrick(lngCounter - 1).Top) Then
                            Exit Sub
                        Else
                            shpBrick(lngCounter - 3).Top = shpBrick(lngCounter - 1).Top
                            shpBrick(lngCounter - 3).Left = shpBrick(lngCounter).Left
                            shpBrick(lngCounter - 2).Top = shpBrick(lngCounter - 1).Top - 405
                            shpBrick(lngCounter - 2).Left = shpBrick(lngCounter - 1).Left
                            shpBrick(lngCounter).Top = shpBrick(lngCounter - 1).Top + 405
                            shpBrick(lngCounter).Left = shpBrick(lngCounter - 1).Left
                            Set shpLeft = shpBrick(lngCounter - 1)
                            Set shpRight = shpBrick(lngCounter - 3)
                            Set shpBottom = shpBrick(lngCounter)
                            lngType = 1
                        End If
                    Case 1
                        If shpLeft.Left = 0 Then
                            
                        ElseIf detProximity("left") And (Not (detObstruction(shpBrick(lngCounter - 1).Left + 405, shpBrick(lngCounter - 1).Top))) Then
                            
                        ElseIf detObstruction(shpBrick(lngCounter - 1).Left, shpBrick(lngCounter - 1).Top) Then
                            Exit Sub
                        Else
                            shpBrick(lngCounter - 3).Top = shpBrick(lngCounter).Top
                            shpBrick(lngCounter - 3).Left = shpBrick(lngCounter - 1).Left
                            shpBrick(lngCounter - 2).Top = shpBrick(lngCounter - 1).Top
                            shpBrick(lngCounter - 2).Left = shpBrick(lngCounter - 1).Left + 405
                            shpBrick(lngCounter).Top = shpBrick(lngCounter - 1).Top
                            shpBrick(lngCounter).Left = shpBrick(lngCounter - 1).Left - 405
                            Set shpLeft = shpBrick(lngCounter)
                            Set shpRight = shpBrick(lngCounter - 2)
                            Set shpBottom = shpBrick(lngCounter - 3)
                            lngType = 2
                        End If
                    Case 2
                        If detObstruction(shpBrick(lngCounter - 1).Left, shpBrick(lngCounter - 1).Top) Then
                            Exit Sub
                        Else
                            shpBrick(lngCounter - 3).Top = shpBrick(lngCounter - 1).Top
                            shpBrick(lngCounter - 3).Left = shpBrick(lngCounter).Left
                            shpBrick(lngCounter - 2).Top = shpBrick(lngCounter - 1).Top + 405
                            shpBrick(lngCounter - 2).Left = shpBrick(lngCounter - 1).Left
                            shpBrick(lngCounter).Top = shpBrick(lngCounter - 1).Top - 405
                            shpBrick(lngCounter).Left = shpBrick(lngCounter - 1).Left
                            Set shpLeft = shpBrick(lngCounter - 3)
                            Set shpRight = shpBrick(lngCounter - 1)
                            Set shpBottom = shpBrick(lngCounter - 2)
                            lngType = 3
                        End If
                    Case 3
                        If shpRight.Left = 3645 Then
                            
                        ElseIf detProximity("right") And (Not (detObstruction(shpBrick(lngCounter - 1).Left - 405, shpBrick(lngCounter - 1).Top))) Then
                            
                        ElseIf detObstruction(shpBrick(lngCounter - 1).Left, shpBrick(lngCounter - 1).Top) Then
                            Exit Sub
                        Else
                            shpBrick(lngCounter - 3).Top = shpBrick(lngCounter).Top
                            shpBrick(lngCounter - 3).Left = shpBrick(lngCounter).Left
                            shpBrick(lngCounter - 2).Top = shpBrick(lngCounter - 1).Top
                            shpBrick(lngCounter - 2).Left = shpBrick(lngCounter - 1).Left - 405
                            shpBrick(lngCounter).Top = shpBrick(lngCounter - 1).Top
                            shpBrick(lngCounter).Left = shpBrick(lngCounter - 1).Left + 405
                            Set shpLeft = shpBrick(lngCounter - 2)
                            Set shpRight = shpBrick(lngCounter)
                            Set shpBottom = shpBrick(lngCounter - 1)
                            lngType = 0
                        End If
                End Select
            Case 6
                Select Case lngType
                    Case 0
                        If shpBottom.Top = 7695 Then
                            Exit Sub
                        ElseIf (shpBrick(lngCounter).Top = 0) Or (shpBrick(lngCounter).Top = 405) Then
                            Exit Sub
                        ElseIf Not (booGrid(shpBrick(lngCounter - 1).Left / 405, shpBrick(lngCounter - 1).Top / 405 - 2) Or booGrid(shpBrick(lngCounter - 1).Left / 405, shpBrick(lngCounter - 1).Top / 405 - 1) Or booGrid(shpBrick(lngCounter - 1).Left / 405, shpBrick(lngCounter - 1).Top / 405 + 1)) Then
                            shpBrick(lngCounter - 3).Left = shpBrick(lngCounter - 1).Left
                            shpBrick(lngCounter - 3).Top = shpBrick(lngCounter - 1).Top - 810
                            shpBrick(lngCounter - 2).Left = shpBrick(lngCounter - 1).Left
                            shpBrick(lngCounter - 2).Top = shpBrick(lngCounter - 1).Top - 405
                            shpBrick(lngCounter).Left = shpBrick(lngCounter - 1).Left
                            shpBrick(lngCounter).Top = shpBrick(lngCounter - 1).Top + 405
                            Set shpLeft = shpBrick(lngCounter - 3)
                            Set shpRight = shpBrick(lngCounter)
                            Set shpBottom = shpBrick(lngCounter)
                            lngType = 1
                        End If
                    Case 1
                        If shpRight.Left = 3645 Then
                            
                        'ElseIf detProximity("right") And (Not (detObstruction(shpBrick(lngCounter - 1).Left - 405, shpBrick(lngCounter - 1).Top))) Then
                            
                        ElseIf shpLeft.Left = 0 Then
                            
                        'ElseIf detProximity("left") And (Not (detObstruction(shpBrick(lngCounter - 1).Left + 405, shpBrick(lngCounter - 1).Top))) Then
                            
                        ElseIf Not (booGrid(shpBrick(lngCounter - 1).Left / 405 - 2, shpBrick(lngCounter - 1).Top / 405) Or booGrid(shpBrick(lngCounter - 1).Left / 405 - 1, shpBrick(lngCounter - 1).Top / 405) Or booGrid(shpBrick(lngCounter - 1).Left / 405 + 1, shpBrick(lngCounter - 1).Top / 405)) Then
                            shpBrick(lngCounter - 3).Left = shpBrick(lngCounter - 1).Left - 810
                            shpBrick(lngCounter - 3).Top = shpBrick(lngCounter - 1).Top
                            shpBrick(lngCounter - 2).Left = shpBrick(lngCounter - 1).Left - 405
                            shpBrick(lngCounter - 2).Top = shpBrick(lngCounter - 1).Top
                            shpBrick(lngCounter).Left = shpBrick(lngCounter - 1).Left + 405
                            shpBrick(lngCounter).Top = shpBrick(lngCounter - 1).Top
                            Set shpLeft = shpBrick(lngCounter - 3)
                            Set shpRight = shpBrick(lngCounter)
                            Set shpBottom = shpBrick(lngCounter)
                            lngType = 0
                        End If
                End Select
        End Select

    ElseIf KeyCode = vbKeyEscape Then
        mnuPause_Click
    End If
End Sub

Private Sub Form_Load()
    lngSpeed = 1
    lngMax = 20
    booIsMoving = False
    lstSpeed.Text = CStr(lngSpeed)
    lngCounter = 1  'don't know why it should be 1...TODO!!
    Dim x, y As Integer
    For x = 0 To 9  'initialize booGrid
        For y = 0 To 19
            booGrid(x, y) = False
        Next
    Next
    lngBrick2 = Int(Rnd * 7)
End Sub

'parse new value of lngSpeed
Private Sub lstSpeed_Click()
    lngSpeed = CLng(lstSpeed.Text)
End Sub

'remember: when frmMain is loaded, the focus is ALWAYS on lstSpeed
Private Sub lstSpeed_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then mnuStart_Click
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuPause_Click()
    timBricks.Enabled = False
    mnuStart.Enabled = True
    lstSpeed.Enabled = True
End Sub

Private Sub mnuStart_Click()
    lstSpeed.Enabled = False
    mnuStart.Enabled = False
    mnuStart.Caption = "Resume (F5)"
    Select Case lngSpeed
        Case 1
            timBricks.Interval = 250
        Case 2
            timBricks.Interval = 200
        Case 3
            timBricks.Interval = 150
        Case 4
            timBricks.Interval = 100
        Case 5
            timBricks.Interval = 50
    End Select
    timBricks.Enabled = True
End Sub

Private Sub timBricks_Timer()
    Dim x, y, i, b As Integer
    lblScore.Caption = CStr(lngScore)
    For i = 0 To 3
        shpPreview(i).Visible = True
    Next
    
    '1. check IsMoving, if no, remove rows AND create a new block
    If booIsMoving = False Then
        'a)check if any row is complete
        For y = 19 To 0 Step -1
            Dim isComplete As Boolean
            isComplete = True
            For x = 0 To 9
                If Not booGrid(x, y) Then isComplete = False
            Next
            If isComplete Then
                'kill an entire row
                For x = 0 To 9
                    Unload shpGrid(x, y)
                Next
                lngScore = lngScore + 100
                'move upper rows down
                For i = (y - 1) To 0 Step -1
                    For x = 0 To 9
                        On Error Resume Next
                        shpGrid(x, i).Top = shpGrid(x, i).Top + 405
                        Set shpGrid(x, i + 1) = shpGrid(x, i)    'reassign shpGrid
                        booGrid(x, i + 1) = booGrid(x, i)   'reassign the values of booGrid
                    Next
                Next
                y = y + 1
            End If
        Next
        
        'b)create another brick
        lngCounter = lngCounter + 4
        Randomize
        lngBrick1 = lngBrick2
        lngBrick2 = Int(Rnd * 7)
        lngType = 0    'initialize the shape
        Select Case lngBrick1
            Case 0  'inverse L
                For i = (lngCounter - 3) To lngCounter
                    Load shpBrick(i)
                Next
                shpBrick(lngCounter - 3).Top = 0: shpBrick(lngCounter - 3).Left = 2025: shpBrick(lngCounter - 3).Visible = True
                shpBrick(lngCounter - 2).Top = 405: shpBrick(lngCounter - 2).Left = 2025: shpBrick(lngCounter - 2).Visible = True
                shpBrick(lngCounter - 1).Top = 810: shpBrick(lngCounter - 1).Left = 1620: shpBrick(lngCounter - 1).Visible = True
                shpBrick(lngCounter).Top = 810: shpBrick(lngCounter).Left = 2025: shpBrick(lngCounter).Visible = True
                Set shpBottom = shpBrick(lngCounter - 1)
                Set shpLeft = shpBrick(lngCounter - 1)
                Set shpRight = shpBrick(lngCounter)
            Case 1  'upright Z
                For i = (lngCounter - 3) To lngCounter
                    Load shpBrick(i)
                Next
                shpBrick(lngCounter - 3).Top = 0: shpBrick(lngCounter - 3).Left = 2025: shpBrick(lngCounter - 3).Visible = True
                shpBrick(lngCounter - 2).Top = 405: shpBrick(lngCounter - 2).Left = 1620: shpBrick(lngCounter - 2).Visible = True
                shpBrick(lngCounter - 1).Top = 405: shpBrick(lngCounter - 1).Left = 2025: shpBrick(lngCounter - 1).Visible = True
                shpBrick(lngCounter).Top = 810: shpBrick(lngCounter).Left = 1620: shpBrick(lngCounter).Visible = True
                Set shpBottom = shpBrick(lngCounter)
                Set shpLeft = shpBrick(lngCounter - 2)
                Set shpRight = shpBrick(lngCounter - 3)
            Case 2  'L
                For i = (lngCounter - 3) To lngCounter
                    Load shpBrick(i)
                Next
                shpBrick(lngCounter - 3).Top = 0: shpBrick(lngCounter - 3).Left = 1620: shpBrick(lngCounter - 3).Visible = True
                shpBrick(lngCounter - 2).Top = 405: shpBrick(lngCounter - 2).Left = 1620: shpBrick(lngCounter - 2).Visible = True
                shpBrick(lngCounter - 1).Top = 810: shpBrick(lngCounter - 1).Left = 1620: shpBrick(lngCounter - 1).Visible = True
                shpBrick(lngCounter).Top = 810: shpBrick(lngCounter).Left = 2025: shpBrick(lngCounter).Visible = True
                Set shpBottom = shpBrick(lngCounter - 1)
                Set shpLeft = shpBrick(lngCounter - 1)
                Set shpRight = shpBrick(lngCounter)
            Case 3  'inverse upright Z
                For i = (lngCounter - 3) To lngCounter
                    Load shpBrick(i)
                Next
                shpBrick(lngCounter - 3).Top = 0: shpBrick(lngCounter - 3).Left = 1620: shpBrick(lngCounter - 3).Visible = True
                shpBrick(lngCounter - 2).Top = 405: shpBrick(lngCounter - 2).Left = 1620: shpBrick(lngCounter - 2).Visible = True
                shpBrick(lngCounter - 1).Top = 405: shpBrick(lngCounter - 1).Left = 2025: shpBrick(lngCounter - 1).Visible = True
                shpBrick(lngCounter).Top = 810: shpBrick(lngCounter).Left = 2025: shpBrick(lngCounter).Visible = True
                Set shpBottom = shpBrick(lngCounter)
                Set shpLeft = shpBrick(lngCounter - 3)
                Set shpRight = shpBrick(lngCounter)
            Case 4  'square
                For i = (lngCounter - 3) To lngCounter
                    Load shpBrick(i)
                Next
                shpBrick(lngCounter - 3).Top = 0: shpBrick(lngCounter - 3).Left = 1620: shpBrick(lngCounter - 3).Visible = True
                shpBrick(lngCounter - 2).Top = 0: shpBrick(lngCounter - 2).Left = 2025: shpBrick(lngCounter - 2).Visible = True
                shpBrick(lngCounter - 1).Top = 405: shpBrick(lngCounter - 1).Left = 1620: shpBrick(lngCounter - 1).Visible = True
                shpBrick(lngCounter).Top = 405: shpBrick(lngCounter).Left = 2025: shpBrick(lngCounter).Visible = True
                Set shpBottom = shpBrick(lngCounter)
                Set shpLeft = shpBrick(lngCounter - 1)
                Set shpRight = shpBrick(lngCounter)
            Case 5  'inverse T
                For i = (lngCounter - 3) To lngCounter
                    Load shpBrick(i)
                Next
                shpBrick(lngCounter - 3).Top = 0: shpBrick(lngCounter - 3).Left = 1620: shpBrick(lngCounter - 3).Visible = True
                shpBrick(lngCounter - 2).Top = 405: shpBrick(lngCounter - 2).Left = 1215: shpBrick(lngCounter - 2).Visible = True
                shpBrick(lngCounter - 1).Top = 405: shpBrick(lngCounter - 1).Left = 1620: shpBrick(lngCounter - 1).Visible = True
                shpBrick(lngCounter).Top = 405: shpBrick(lngCounter).Left = 2025: shpBrick(lngCounter).Visible = True
                Set shpBottom = shpBrick(lngCounter - 1)
                Set shpLeft = shpBrick(lngCounter - 2)
                Set shpRight = shpBrick(lngCounter)
            Case 6  'line
                For i = (lngCounter - 3) To lngCounter
                    Load shpBrick(i)
                Next
                shpBrick(lngCounter - 3).Top = 0: shpBrick(lngCounter - 3).Left = 1215: shpBrick(lngCounter - 3).Visible = True
                shpBrick(lngCounter - 2).Top = 0: shpBrick(lngCounter - 2).Left = 1620: shpBrick(lngCounter - 2).Visible = True
                shpBrick(lngCounter - 1).Top = 0: shpBrick(lngCounter - 1).Left = 2025: shpBrick(lngCounter - 1).Visible = True
                shpBrick(lngCounter).Top = 0: shpBrick(lngCounter).Left = 2430: shpBrick(lngCounter).Visible = True
                Set shpBottom = shpBrick(lngCounter - 3)
                Set shpLeft = shpBrick(lngCounter - 3)
                Set shpRight = shpBrick(lngCounter)
        End Select
        
        'c)set preview
        Select Case lngBrick2
            Case 0
                shpPreview(0).Top = 1245: shpPreview(0).Left = 525
                shpPreview(1).Top = 1245: shpPreview(1).Left = 930
                shpPreview(2).Top = 840: shpPreview(2).Left = 930
                shpPreview(3).Top = 435: shpPreview(3).Left = 930
            Case 1
                shpPreview(0).Top = 1245: shpPreview(0).Left = 525
                shpPreview(1).Top = 840: shpPreview(1).Left = 525
                shpPreview(2).Top = 840: shpPreview(2).Left = 930
                shpPreview(3).Top = 435: shpPreview(3).Left = 930
            Case 2
                shpPreview(0).Top = 1245: shpPreview(0).Left = 525
                shpPreview(1).Top = 840: shpPreview(1).Left = 525
                shpPreview(2).Top = 435: shpPreview(2).Left = 525
                shpPreview(3).Top = 1245: shpPreview(3).Left = 930
            Case 3
                shpPreview(0).Top = 840: shpPreview(0).Left = 525
                shpPreview(1).Top = 435: shpPreview(1).Left = 525
                shpPreview(2).Top = 840: shpPreview(2).Left = 930
                shpPreview(3).Top = 1245: shpPreview(3).Left = 930
            Case 4
                shpPreview(0).Top = 1245: shpPreview(0).Left = 525
                shpPreview(1).Top = 840: shpPreview(1).Left = 525
                shpPreview(2).Top = 840: shpPreview(2).Left = 930
                shpPreview(3).Top = 1245: shpPreview(3).Left = 930
            Case 5
                shpPreview(0).Top = 1245: shpPreview(0).Left = 525
                shpPreview(1).Top = 1245: shpPreview(1).Left = 930
                shpPreview(2).Top = 1245: shpPreview(2).Left = 1335
                shpPreview(3).Top = 840: shpPreview(3).Left = 930
            Case 6
                shpPreview(0).Top = 840: shpPreview(0).Left = 120
                shpPreview(1).Top = 840: shpPreview(1).Left = 525
                shpPreview(2).Top = 840: shpPreview(2).Left = 930
                shpPreview(3).Top = 840: shpPreview(3).Left = 1335
        End Select
        
        booIsMoving = True
        Exit Sub
    End If
    
    '2. check if it reached the bottom brink
    If shpBottom.Top = 7695 Then
        For i = (lngCounter - 3) To lngCounter
            x = shpBrick(i).Left / 405
            y = shpBrick(i).Top / 405
            booGrid(x, y) = True
            Set shpGrid(x, y) = shpBrick(i)
            If y < lngMax Then lngMax = y
        Next
        booIsMoving = False
        Exit Sub
    End If
    
    '3. check if the lowest part of the block has touched other blocks
    For b = (lngCounter - 3) To lngCounter
        If booGrid(shpBrick(b).Left / 405, (shpBrick(b).Top + 405) / 405) Then
            For i = (lngCounter - 3) To lngCounter
                x = shpBrick(i).Left / 405
                y = shpBrick(i).Top / 405
                booGrid(x, y) = True
                Set shpGrid(x, y) = shpBrick(i)
                If y < lngMax Then lngMax = y
                
                'check if gameover
                If lngMax <= 0 Then
                    'gameover!!!!
                    MsgBox "GAME OVER!", vbOKOnly, "Game Over"
                    End
                End If
                
            Next
            booIsMoving = False
            Exit Sub
        End If
    Next
    
    '4. move the block further down
    shpBrick(lngCounter).Top = shpBrick(lngCounter).Top + 405
    shpBrick(lngCounter - 1).Top = shpBrick(lngCounter - 1).Top + 405
    shpBrick(lngCounter - 2).Top = shpBrick(lngCounter - 2).Top + 405
    shpBrick(lngCounter - 3).Top = shpBrick(lngCounter - 3).Top + 405
End Sub

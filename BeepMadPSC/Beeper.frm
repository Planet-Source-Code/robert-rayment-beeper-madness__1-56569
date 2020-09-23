VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  Beeper Madness"
   ClientHeight    =   555
   ClientLeft      =   150
   ClientTop       =   0
   ClientWidth     =   3855
   DrawWidth       =   2
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   37
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   257
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFile 
      Caption         =   "Start"
      Height          =   345
      Index           =   1
      Left            =   960
      TabIndex        =   3
      Top             =   90
      Width           =   900
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "Exit"
      Height          =   345
      Index           =   3
      Left            =   2850
      TabIndex        =   2
      Top             =   90
      Width           =   900
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "Stop"
      Height          =   345
      Index           =   2
      Left            =   1905
      TabIndex        =   1
      Top             =   90
      Width           =   900
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "File"
      Height          =   345
      Index           =   0
      Left            =   30
      TabIndex        =   0
      Top             =   90
      Width           =   900
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' BeeperMadness by Robert Rayment  Oct 2004
' Plays spk files
' 160 spk files
' VB6 InstrRev

Option Explicit
Option Base 1

Private Declare Function Beep Lib "Kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
' NB Not Win98

Private aSTARTSW As Boolean
Private PathSpec$, SpkPath$, FileSpec$

Private Dura() As Long, Freq() As Long
Private Num As Long  ' Num notes
Private TDuration As Long
Private TScale As Long

Private CommonDialog1 As OSDialog

Private Sub Form_Load()
   FileSpec$ = ""
   PathSpec$ = App.Path
   If Right$(PathSpec$, 1) <> "\" Then PathSpec$ = PathSpec$ & "\"
   SpkPath$ = PathSpec$ & "SPK\"
   If Not PathExists(SpkPath$) Then
      MsgBox " Can't find SPK files", vbCritical, "Beeper"
      Unload Me
      End
   End If
End Sub

Private Sub cmdFile_Click(Index As Integer)
   Select Case Index
   Case 0   ' Get SPK file & play
      aSTARTSW = True
      Caption = ""
      GetSPK
      If LenB(FileSpec$) <> 0 Then
         TranslateSPK FileSpec$
         Caption = GetFileName(FileSpec$)
      Else
         MsgBox " Funny SPK file", vbCritical, "Beeper"
      End If
   Case 1   'Start
      PlaySPK
   Case 2   ' Stop
      aSTARTSW = True
   Case 3   ' Exit
      aSTARTSW = True
      Unload Me
      End
   End Select
End Sub

Private Sub GetSPK()
Dim Title$, Filt$, InDir$
Dim FIndex As Long
   Set CommonDialog1 = New OSDialog
   Title$ = "Open Beep file"
   Filt$ = "Open Beep (*.spk)|*.spk"
   InDir$ = SpkPath$
   CommonDialog1.ShowOpen FileSpec$, Title$, Filt$, InDir$, "", Me.hWnd, FIndex
   ' FIndex = 1 *.spk file
   Set CommonDialog1 = Nothing
End Sub

Private Sub TranslateSPK(SPKSpec$)
' Public Num
Dim A$
   On Error Resume Next
   TDuration = 5
   TScale = 32767
   ReDim Freq(1), Dura(1)
   Num = 0
   Open SPKSpec$ For Input As #1
   Do Until EOF(1)
      Line Input #1, A$
      TranslateLine A$
   Loop
   Close
End Sub

Private Sub TranslateLine(A$)
' Public Num
Dim E$, V$
Dim p As Long
Dim TFreq As Long
Dim PFreq As Long
Dim Tempo As Long
Dim Gap As Long
   On Error Resume Next
   PFreq = 1000
   TFreq = 1000
   
   If A$ = "" Then Exit Sub
   For p = 1 To Len(A$)
      E$ = Mid$(A$, p, 1)
      If Asc(E$) < 64 Then    'Number, build value
         If E$ <> " " Then V$ = V$ + E$
      Else
         Select Case E$
         Case " ", "P", "R", "B", "E"
         Case "F", "f"
            If V$ <> "" Then
               TFreq = Val(V$)
            Else
               TFreq = PFreq    'Prev freq
            End If
            
            If TFreq = 0 Then
               TFreq = PFreq
               TDuration = 0
            End If
            
            PFreq = TFreq
            
            Num = Num + 1
            ReDim Preserve Freq(Num), Dura(Num)
            Freq(Num) = PFreq
            Dura(Num) = TDuration
            If Gap <> 0 Then
               Num = Num + 1
               ReDim Preserve Freq(Num), Dura(Num)
               Freq(Num) = 32767&
               Dura(Num) = Gap
            End If
         Case "D", "d"
            TDuration = Val(V$) * TScale
         Case "G", "g"
            Gap = Val(V$) * TScale
         Case "T", "t"
            Tempo = Val(V$)
            TScale = TScale * Tempo / 256
         End Select
         If Num > 0 Then
            If Freq(Num) < 37 Then Freq(Num) = 37
         End If
         V$ = ""
      End If
   Next p
End Sub

Private Sub PlaySPK()
' Public Num
Dim i As Long
Dim Frequency As Long
Dim Length As Long
   If Num < 1 Then Exit Sub
   
   aSTARTSW = False
   
   For i = 1 To Num
      DoEvents
      If aSTARTSW Then Exit For
      Length = Dura(i) / Freq(i)
      Frequency = Freq(i)
      Beep Frequency, Length
   Next i
End Sub

Public Function PathExists(ByVal InPathName$) As Boolean
   On Error Resume Next
   PathExists = (Dir$(InPathName$ & "\nul") <> "")
End Function

Public Function GetFileName(FSpec$) As String
Dim p As Long
   GetFileName = FSpec$
   p = InStrRev(FSpec$, "\")
   If p <> 0 Then
      GetFileName = Right$(FSpec$, Len(FSpec$) - p)
   End If
End Function

Private Sub Form_Unload(Cancel As Integer)
   aSTARTSW = False
   Unload Me
   End
End Sub

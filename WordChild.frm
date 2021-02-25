VERSION 5.00
Begin VB.Form WordChild 
   ClientHeight    =   4605
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   6420
   Icon            =   "WordChild.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4605
   ScaleWidth      =   6420
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Height          =   4575
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   6375
   End
   Begin VB.Menu mnuArkiv 
      Caption         =   "&Arkiv"
      Begin VB.Menu mnuNytt 
         Caption         =   "Nytt"
      End
      Begin VB.Menu mnu÷ppna 
         Caption         =   "÷ppna"
      End
      Begin VB.Menu mnuSpara 
         Caption         =   "Spara "
      End
      Begin VB.Menu mnuSparaSom 
         Caption         =   "Spara Som"
      End
      Begin VB.Menu mnuSkriv 
         Caption         =   "Skriv ut"
      End
      Begin VB.Menu mnuSt‰ng 
         Caption         =   "St‰ng"
      End
      Begin VB.Menu mnuSkilj 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAvsluta 
         Caption         =   "Avsluta"
      End
   End
   Begin VB.Menu mnuRedigera 
      Caption         =   "&Regdigera"
      Begin VB.Menu mnu≈ngra 
         Caption         =   "≈ngra"
      End
      Begin VB.Menu mnuSkilj2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuKlipp 
         Caption         =   "Klipp ut"
      End
      Begin VB.Menu mnuKopiera 
         Caption         =   "Kopiera"
      End
      Begin VB.Menu mnuKlistra 
         Caption         =   "Klista in"
      End
      Begin VB.Menu mnuTaBort 
         Caption         =   "Ta bort"
      End
      Begin VB.Menu mnuSkilj3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMarkera 
         Caption         =   "Markera Allt"
      End
   End
   Begin VB.Menu mnutecken 
      Caption         =   "&Tecken"
      Begin VB.Menu mnuStorlek 
         Caption         =   "Teckenstorlek"
         Begin VB.Menu mnu8 
            Caption         =   "8"
         End
         Begin VB.Menu mnu10 
            Caption         =   "10"
         End
         Begin VB.Menu mnu12 
            Caption         =   "12"
         End
         Begin VB.Menu mnu14 
            Caption         =   "14"
         End
         Begin VB.Menu mnu18 
            Caption         =   "18"
         End
         Begin VB.Menu mnu24 
            Caption         =   "24"
         End
         Begin VB.Menu mnu30 
            Caption         =   "30"
         End
         Begin VB.Menu mnu48 
            Caption         =   "48"
         End
         Begin VB.Menu mnu72 
            Caption         =   "72"
         End
      End
      Begin VB.Menu mnuSnitt 
         Caption         =   "Teckensnitt"
         Begin VB.Menu mnuArial 
            Caption         =   "Arial"
         End
         Begin VB.Menu mnuComic 
            Caption         =   "Comic Sans MS"
         End
         Begin VB.Menu mnuCourier 
            Caption         =   "Courier"
         End
         Begin VB.Menu mnuGaramond 
            Caption         =   "Garamond"
         End
         Begin VB.Menu mnuItalic 
            Caption         =   "Italic"
         End
         Begin VB.Menu mnuSystem 
            Caption         =   "System"
         End
         Begin VB.Menu mnuTahoma 
            Caption         =   "Tahoma"
         End
         Begin VB.Menu mnuTimes 
            Caption         =   "Times New Roman"
         End
         Begin VB.Menu mnuVerdana 
            Caption         =   "Verdana"
         End
      End
   End
   Begin VB.Menu mnuVisa 
      Caption         =   "&Visa"
      Begin VB.Menu mnuKnapprad 
         Caption         =   "Knapprad"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuFˆnster 
      Caption         =   "&Fˆnster"
      WindowList      =   -1  'True
      Begin VB.Menu mnu÷verlappade 
         Caption         =   "÷verlappade"
      End
      Begin VB.Menu mnuSida 
         Caption         =   "Sida vid sida"
      End
      Begin VB.Menu mnu÷verUnder 
         Caption         =   "÷ver och under"
      End
   End
   Begin VB.Menu mnufrÂga 
      Caption         =   "?"
      Begin VB.Menu mnuOm 
         Caption         =   "Om"
      End
   End
End
Attribute VB_Name = "WordChild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub PrintText(text As String)
Dim i As Integer, j As Integer, ord As String
Dim MarginLeft As Double, MarginTop As Double
  

MarginLeft = 1.91
MarginTop = 2.54
Printer.CurrentX = MarginLeft
Printer.CurrentY = MarginTop
On Error GoTo errh:
Printer.Print
i = 1

Do Until 1 > Len(text)

ord = ""
ord = ord & Mid$(text, i, 1)
i = i + 1
Loop
'FÂr det plats pÂ raden
If (Printer.CurrentX + Printer.TextWidth(ord)) > Printer.ScaleWidth Then
'Om inte n‰sta rad matas fram
Printer.Print
Printer.CurrentX = MarginLeft
End If
'Skicka ett ord till skrivaren
Printer.Print ord;
'Ta hand om mellanslaget eller eventuella styrtecken mellan orden
Do Until i > Len(text) Or Mid$(text, i, 1) > " "
Select Case Mid$(text, i, 1)
 Case " "     'mellanslag
  Printer.Print " ";
 Case Chr$(10) 'Ny rad
  Printer.Print
  Printer.CurrentX = MarginLeft
 Case Chr$(9)    'Tab
  Printer.Print Tab
End Select
  i = i + 1
 Loop

'Avsluta utskriften
Printer.EndDoc
Exit Sub
errh:
If Err = 482 Then
  MsgBox "Kontrollera att skrivaren ‰r ordenligt anslutan och pÂslagen", 48, "SuperWriter"
End If
Exit Sub
End Sub
Public Sub PrintFile()
Dim TextArea As String
Dim flaggor As Integer
  'V‰lj skrivare
  SuperWriter.CommonDialog1.DialogTitle = "Skriv ut"
  SuperWriter.CommonDialog1.CancelError = True
  flaggor = cdlPDNoPageNums
   If SuperWriter.ActiveForm.Text1.SelLength = 0 Then
   flaggor = flaggor + cdlPDAllPages
   Else
   flaggor = flaggor + cdlPDSelection
   End If
   SuperWriter.CommonDialog1.Flags = flaggor
   
   'Aktiver felhantering
   On Error GoTo PrintErrorhandler:
   SuperWriter.CommonDialog1.ShowPrinter
   
   TextArea = SuperWriter.ActiveForm.Text1.text
   PrintText TextArea
   Exit Sub
   
PrintErrorhandler:
    If Err = CDCANCEL Then
    Exit Sub
    Else
    Resume Next
    End If
   End Sub
Public Sub FileSave(sFilnamn As String)
Dim i As Integer
i = FreeFile

'Aktivera felhantering
On Error GoTo Errhandler:

'Pang pÂ och spara
Open sFilnamn For Output As i
 Print #1, SuperWriter.ActiveForm.Text1.text
 Close i
 'Uppdatera formul‰ret
 SuperWriter.ActiveForm.Tag = SAVED
 SuperWriter.ActiveForm.Caption = sFilnamn
 SuperWriter.StatusBar.Panels(1).text = "Filstatus SAVED"
 Exit Sub
Errhandler:
   MsgBox "Ett fel har uppstÂtt. Filen sparades inte", 48, "SuperWriter"
End Sub
Public Sub EditCopy()
SuperWriter.ActiveForm.SetFocus
SendKeys "^c", True
End Sub
Public Sub EditCut()
SuperWriter.ActiveForm.SetFocus
SendKeys "^x", True
End Sub
Public Sub EditPaste()
SuperWriter.ActiveForm.SetFocus
SendKeys "^v", True
End Sub
Public Sub FileOpen()
Dim NewFile As New WordChild
Dim sFilnamn As String
Dim i As Integer
  i = FreeFile
   
   'Aktivera felhantering
   On Error GoTo errorhandler:
   
   SuperWriter.CommonDialog1.DialogTitle = "÷ppna"
   SuperWriter.CommonDialog1.CancelError = True
   SuperWriter.CommonDialog1.Filter = "Text Format|*.txt"
   SuperWriter.CommonDialog1.ShowOpen
   sFilnamn = SuperWriter.CommonDialog1.FileName
  

'L‰s in den utvalda filen
Open sFilnamn For Input As i
  NewFile.Text1.text = Input(LOF(i), i)
  Close i
  
  'Ordna till det nya formul‰ret
  NewFile.Caption = sFilenamn
  NewFile.Tag = SAVED
  SuperWriter.StatusBar.Panels(1).text = "Filstatus SAVED"
  Exit Sub
errorhandler:
  If Not (Err.Number = CDCANCEL) Then MsgBox "Filen finns inte", 48, "SuperWriter"
  Unload NewFile
End Sub
Public Sub FileSaveAs()
Dim sFilnamn As String

'Aktivera felhantering
On Error GoTo Errhandler:

SuperWriter.CommonDialog1.Filter = "Text Format|*.txt"
SuperWriter.CommonDialog1.DefaultExt = ".txt"
SuperWriter.CommonDialog1.DialogTitle = "Spara Som"
SuperWriter.CommonDialog1.ShowSave
sFilnamn = SuperWriter.CommonDialog1.FileName
'Spara
FileSave sFilnamn
Exit Sub
Errhandler:
  If Not (Err.Number = CDCANCEL) Then MsgBox "Ett fel har uppstÂtt. Filen sparades inte", 48, "Superwriter"
End Sub

Public Sub NewFile()
Dim NewDoc As New WordChild
NewDoc.Show
End Sub
Private Sub Form_Activate()
If Me.Tag = SAVED Then
  SuperWriter.StatusBar.Panels(1).text = "Filstatus SAVED"
  Else
  SuperWriter.StatusBar.Panels(1).text = "Filstatus UNSAVED"
  End If
End Sub

Private Sub Form_Load()
  Me.Caption = INITFILNAMN
  Me.Tag = UNSAVED
  If Forms.Count = 2 Then SuperWriter.Knapprad÷ppna
End Sub

Private Sub Form_Resize()
 Text1.Height = ScaleHeight
 Text1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim iSvar As Integer
  Cancel = False
  If Me.Tag = UNSAVED Then
  iSvar = MsgBox("Filen " & Me.Caption & " har ‰ndrats. Vill du spara den?", 3, "SuperWriter")
  Select Case iSvar
    Case 6
     If Me.Caption = INITFILNAMN Then
       FileSaveAs
     Else
       FileSave Me.Caption
     End If
    Case 2
     Cancel = True
End Select
End If
If Forms.Count = 2 Then
  SuperWriter.KnappradInit
  SuperWriter.StatusBar.Panels(1).text = "Filstatus"
  End If
End Sub

Private Sub mnu10_Click()
Text1.FontSize = 10
End Sub

Private Sub mnu12_Click()
Text1.FontSize = 12
End Sub

Private Sub mnu14_Click()
Text1.FontSize = 14
End Sub

Private Sub mnu18_Click()
Text1.FontSize = 18
End Sub

Private Sub mnu24_Click()
Text1.FontSize = 24
End Sub

Private Sub mnu30_Click()
Text1.FontSize = 30
End Sub

Private Sub mnu48_Click()
Text1.FontSize = 48
End Sub

Private Sub mnu72_Click()
Text1.FontSize = 72
End Sub

Private Sub mnu8_Click()
Text1.FontSize = 8
End Sub

Private Sub mnuArial_Click()
Text1.Font = "Arial"
End Sub

Private Sub mnuAvsluta_Click()
End
End Sub

Private Sub mnuComic_Click()
Text1.Font = "Comic Sans MS"
End Sub

Private Sub mnuCourier_Click()
Text1.Font = "Courier"
End Sub

Private Sub mnuGaramond_Click()
Text1.Font = "Garamond"
End Sub

Private Sub mnuItalic_Click()
Text1.Font = "Italic"
End Sub

Private Sub mnuKlipp_Click()
EditCut
End Sub

Private Sub mnuKlistra_Click()
EditPaste
End Sub

Private Sub mnuKnapprad_Click()
mnuKnapprad.Checked = Not (mnuKnapprad.Checked)
SuperWriter.Knapprad.Visible = mnuKnapprad.Checked
End Sub

Private Sub mnuKopiera_Click()
EditCopy
End Sub

Private Sub mnuMarkera_Click()
SuperWriter.ActiveForm.Text1.SelStart = 0
Text1.SelLength = Len(Text1)
End Sub

Private Sub mnuNytt_Click()
NewFile
End Sub

Private Sub mnuOm_Click()
frmOm.Show 1
End Sub

Private Sub mnuSida_Click()
SuperWriter.Arrange TILE_VERTICAL
End Sub

Private Sub mnuSkriv_Click()
PrintFile
End Sub

Private Sub mnuSpara_Click()
If Me.Caption = INITFILNAMN Then
  FileSaveAs
  Else
  FileSave Me.Caption
  End If
End Sub

Private Sub mnuSparaSom_Click()
FileSaveAs
End Sub

Private Sub mnuSt‰ng_Click()
Unload Me
End Sub

Private Sub mnuSystem_Click()
Text1.Font = "System"
End Sub

Private Sub mnuTaBort_Click()
SuperWriter.ActiveForm.SetFocus
SendKeys "{del}"
End Sub

Private Sub mnuTahoma_Click()
Text1.Font = "Tahoma"
End Sub

Private Sub mnuTimes_Click()
Text1.Font = "Times New Roman"
End Sub

Private Sub mnuVerdana_Click()
Text1.Font = "Verdana"
End Sub

Private Sub mnu≈ngra_Click()
SuperWriter.ActiveForm.SetFocus
SendKeys "^z", True
End Sub

Private Sub mnu÷ppna_Click()
FileOpen
End Sub

Private Sub mnu÷verlappade_Click()
SuperWriter.Arrange CASCADE
End Sub

Private Sub mnu÷verUnder_Click()
SuperWriter.Arrange TILE_HORIZONTIAL
End Sub

Private Sub Text1_Change()
  If Not Me.Tag = UNSAVED Then
  Me.Tag = UNSAVED
  SuperWriter.StatusBar.Panels(1).text = "Filstatus UNSAVED"
  End If
End Sub

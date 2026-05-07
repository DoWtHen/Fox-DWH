VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "WO Infók"
   ClientHeight    =   6740
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4210
   OleObjectBlob   =   "kinezett word 3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Modul tetején:
Public LastWO As String
Public LastTopBot As String
Public LastKocsiSzam As String

Private Sub UserForm_Initialize()
    cmbTopBot.Clear
    cmbTopBot.AddItem "TOP"
    cmbTopBot.AddItem "BOT"

    txtWO.Text = LastWO
    cmbTopBot.Text = LastTopBot
    txtKocsiSzam.Text = LastKocsiSzam
End Sub

Private Sub cmdOK_Click()

    Dim WO As String
    Dim TopBot As String
    Dim KocsiSzam As String
    Dim adat As Variant

    adat = Array(10000, 200000, 30000, 40000, 50000, 60000, 70000, 80000, 90000, 100000)

    WO = txtWO.Text
    TopBot = cmbTopBot.Text
    KocsiSzam = txtKocsiSzam.Text

    If WO = "" Or TopBot = "" Or KocsiSzam = "" Then
        MsgBox "Minden mezõt ki kell tölteni!", vbExclamation
        Exit Sub
    End If

    ' Értékek megjegyzése
    LastWO = WO
    LastTopBot = TopBot
    LastKocsiSzam = KocsiSzam

    ' Szöveg és formázás törlése
    With ActiveDocument.Content
        .Font.Reset
        .ParagraphFormat.Reset
        .Text = ""
    End With

    ' Új tartalom
    Selection.TypeText "WO:  " & WO & vbCrLf
    Selection.TypeText TopBot & vbTab & KocsiSzam


    'FORMÁZÁS lépései
    Selection.WholeStory
    Selection.Font.Bold = wdToggle
    Selection.Font.Size = 100
    Selection.PageSetup.Orientation = wdOrientLandscape
    Selection.WholeStory
    With ActiveDocument.Styles(wdStyleNormal).Font
        If .NameFarEast = .NameAscii Then
            .NameAscii = ""
        End If
        .NameFarEast = ""
    End With
    With ActiveDocument.PageSetup
        .VerticalAlignment = wdAlignVerticalCenter
        .BookFoldPrintingSheets = 1
        .GutterPos = wdGutterPosLeft
    End With
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    
    
   ' Unload Me

End Sub


Private Sub CmdNyomtat_Click()
    ActiveDocument.PrintOut
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub


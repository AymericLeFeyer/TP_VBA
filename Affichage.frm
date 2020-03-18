VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Affichage 
   Caption         =   "Affichage des étudiants"
   ClientHeight    =   6612
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4080
   OleObjectBlob   =   "Affichage.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Affichage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Result() As String


Private Sub CommandButton1_Click()
Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Etatcivil.txt" For Input As #1

ListBox1.Clear


While Not EOF(1)
    Line Input #1, ContenuLigne
    Result = Split(ContenuLigne, ",")
    ListBox1.AddItem "Etudiant " & Result(0) & " : " & Result(1) & " " & Result(2) & ", " & Result(3)
Wend

Close #1

End Sub

Private Sub ListBox1_Click()

End Sub

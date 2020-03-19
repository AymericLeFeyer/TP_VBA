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
'Ouverture du fichier Etatcivil
Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Etatcivil.txt" For Input As #1
'On clear la liste pour ne pas la surcharger
ListBox1.Clear
'On parcours le fichier jusqu'au bout
While Not EOF(1)
    'On recupere la ligne suivante
    Line Input #1, ContenuLigne
    'On separe les valeurs par la virgule, dans le tableau Result
    Result = Split(ContenuLigne, ",")
    'On ajoute a la liste les informations de maniere un peu plus propre
    ListBox1.AddItem "Etudiant " & Result(0) & " : " & Result(1) & " " & Result(2) & ", " & Result(3)
Wend
'On ferme le fichier
Close #1

End Sub

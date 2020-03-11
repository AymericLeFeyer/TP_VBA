VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Creation 
   Caption         =   "Création d'un état civil"
   ClientHeight    =   6072
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3924
   OleObjectBlob   =   "Creation.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Creation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim File As String, num As String, nom As String, prenom As String, natio As String

Private Sub CommandButton1_Click()
num = TextBox1.Value
nom = TextBox2.Value
prenom = TextBox3.Value
natio = TextBox4.Value

If (num <> "" And nom <> "" And prenom <> "" And natio <> "") Then
    If (NbOc(num, ",") = 0 And NbOc(nom, ",") = 0 And NbOc(prenom, ",") = 0 And NbOc(natio, ",") = 0) Then
        Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Etatcivil.txt" For Append As #1
        Print #1, num & "," & nom & "," & prenom & "," & natio
        Close #1
        MsgBox "L'étudiant est créé"
    Else
        MsgBox "La virgule est un caractère interdit !"
    End If
Else
    MsgBox "Au moins un champ est vide !"
End If



End Sub

Function NbOc(Chaine As String, Ch As String, Optional RC As Boolean = False) As Long
    If RC Then
        NbOc = (Len(Chaine) - Len(Replace(Chaine, Ch, "", , , 0))) / Len(Ch)
    Else
        NbOc = (Len(Chaine) - Len(Replace(Chaine, Ch, "", , , 1))) / Len(Ch)
    End If
End Function




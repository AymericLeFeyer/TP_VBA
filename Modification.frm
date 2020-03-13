VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Modification 
   Caption         =   "Modification des fichiers"
   ClientHeight    =   9600.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5892
   OleObjectBlob   =   "Modification.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Modification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim numEtudiant, num, nom, prenom, natio As String
Dim TP, DS As String
Dim CP, Adresse As String
Dim etudiantExist As Boolean
Dim Result() As String



Private Sub CommandButton1_Click()
    etudiantExist = False
    
    Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Etatcivil.txt" For Input As #1
    
    While Not EOF(1)
        Line Input #1, ContenuLigneEC
        Result = Split(ContenuLigneEC, ",")
        If (Result(0) = numEtudiant) Then
            etudiantExist = True
            num = Result(0)
            TextBox2.Enabled = True
            TextBox2.Value = num
            nom = Result(1)
            TextBox9.Enabled = True
            TextBox9.Value = nom
            prenom = Result(2)
            TextBox3.Enabled = True
            TextBox3.Value = prenom
            natio = Result(3)
            TextBox4.Enabled = True
            TextBox4.Value = natio
        End If
    Wend
    
    Close #1
    
    If (etudiantExist = True) Then
        Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Notes.txt" For Input As #2
        Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Adresses.txt" For Input As #3
        While Not EOF(2)
            Line Input #2, ContenuLigneN
            Result = Split(ContenuLigneN, ",")
            If (Result(0) = numEtudiant) Then
                TP = Result(1)
                TextBox6.Enabled = True
                TextBox6.Value = TP
                DS = Result(2)
                TextBox5.Enabled = True
                TextBox5.Value = DS
            End If
        Wend
        Close #2
        While Not EOF(3)
            Line Input #3, ContenuLigneA
            Result = Split(ContenuLigneA, ",")
            If (Result(0) = numEtudiant) Then
                CP = Result(1)
                TextBox8.Enabled = True
                TextBox8.Value = CP
                Adresse = Result(2)
                TextBox7.Enabled = True
                TextBox7.Value = Adresse
            End If
        Wend
        Close #3
    
    Else
        MsgBox "Cet étudiant n'existe pas"
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        TextBox6.Enabled = False
        TextBox7.Enabled = False
        TextBox8.Enabled = False
        TextBox9.Enabled = False
        TextBox2.Value = ""
        TextBox3.Value = ""
        TextBox4.Value = ""
        TextBox5.Value = ""
        TextBox6.Value = ""
        TextBox7.Value = ""
        TextBox8.Value = ""
        TextBox9.Value = ""
        
        
    End If
    
    

End Sub

Private Sub CommandButton2_Click()

    numEtudiant = TextBox1.Value
    num = TextBox2.Value
    nom = TextBox9.Value
    prenom = TextBox3.Value
    natio = TextBox4.Value
    TP = TextBox6.Value
    DS = TextBox5.Value
    CP = TextBox8.Value
    Adresse = TextBox7.Value

End Sub

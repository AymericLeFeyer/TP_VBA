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
Dim Tablo(10000) As String

Private Sub CheckBox1_Click()
    If (CheckBox1.Value = True) Then
        If (CheckBox2.Enabled = True) Then
            CheckBox2.Value = True
        End If
        If (CheckBox3.Enabled = True) Then
            CheckBox3.Value = True
        End If
    End If
End Sub

Private Sub CommandButton1_Click()
    CheckBox1.Value = False
    CheckBox2.Value = False
    CheckBox3.Value = False

    numEtudiant = TextBox1.Value
    
    etudiantExist = False
    AlreadyAdd = False
    
    Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Etatcivil.txt" For Input As #1
    
    While Not EOF(1)
        Line Input #1, ContenuLigneEC
        Result = Split(ContenuLigneEC, ",")
        If (Result(0) = numEtudiant) Then
            etudiantExist = True
            num = Result(0)
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
            CheckBox1.Enabled = True
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
                AlreadyAdd = True
                CheckBox2.Enabled = True
            Else
                If (AlreadyAdd = False) Then
                    TextBox6.Enabled = False
                    TextBox6.Value = ""
                    TextBox5.Enabled = False
                    TextBox5.Value = ""
                End If
            End If
        Wend
        
        Close #2
        AlreadyAdd = False
        
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
                AlreadyAdd = True
                CheckBox3.Enabled = True
            Else
                If (AlreadyAdd = False) Then
                    TextBox8.Enabled = False
                    TextBox8.Value = ""
                    TextBox7.Enabled = False
                    TextBox7.Value = ""
                End If
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
        CheckBox1.Enabled = False
        CheckBox2.Enabled = False
        CheckBox3.Enabled = False
        CheckBox1.Value = False
        CheckBox2.Value = False
        CheckBox3.Value = False
        
            
    End If

End Sub

Private Sub CommandButton2_Click()

    numEtudiant = TextBox1.Value
    num = TextBox2.Value
    nom = TextBox9.Value
    prenom = TextBox3.Value
    natio = TextBox4.Value
    If (IsNumeric(TextBox6.Value)) Then
        TP = TextBox6.Value
    Else
        MsgBox "La note de TP est invalide"
    End If
    If (IsNumeric(TextBox5.Value)) Then
        DS = TextBox5.Value
    Else
        MsgBox "La note de DS est invalide"
    End If
    CP = TextBox8.Value
    Adresse = TextBox7.Value
    
    'Recup etatcivil.txt
    i = 0
    Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Etatcivil.txt" For Input As #1
    While Not EOF(1)
        Line Input #1, ContenuLigne
        Result = Split(ContenuLigne, ",")
        If Result(0) = numEtudiant Then
            Result(1) = nom
            Result(2) = prenom
            Result(3) = natio
        End If
        Tablo(i) = Result(0) & "," & Result(1) & "," & Result(2) & "," & Result(3)
        i = i + 1
    Wend
    Close #1
    Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Etatcivil.txt" For Output As #1
    j = 0
    While j < i
        Print #1, Tablo(j)
        j = j + 1
    Wend
    Close #1
    
    'Recup notes.txt
    i = 0
    Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Notes.txt" For Input As #1
    While Not EOF(1)
        Line Input #1, ContenuLigne
        Result = Split(ContenuLigne, ",")
        If Result(0) = numEtudiant Then
            Result(1) = TP
            Result(2) = DS
        End If
        Tablo(i) = Result(0) & "," & Result(1) & "," & Result(2)
        i = i + 1
    Wend
    Close #1
    Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Notes.txt" For Output As #1
    j = 0
    While j < i
        Print #1, Tablo(j)
        j = j + 1
    Wend
    Close #1
    
    'Recup adresses.txt
    i = 0
    Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Adresses.txt" For Input As #1
    While Not EOF(1)
        Line Input #1, ContenuLigne
        Result = Split(ContenuLigne, ",")
        If Result(0) = numEtudiant Then
            Result(1) = CP
            Result(2) = Adresse
        End If
        Tablo(i) = Result(0) & "," & Result(1) & "," & Result(2)
        i = i + 1
    Wend
    Close #1
    Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Adresses.txt" For Output As #1
    j = 0
    While j < i
        Print #1, Tablo(j)
        j = j + 1
    Wend
    Close #1
    
    MsgBox "Etudiant modifié"
    
    
    

End Sub

Private Sub CommandButton3_Click()
    If (CheckBox1.Value = False And CheckBox2.Value = False And CheckBox3.Value = False) Then
        MsgBox "Aucune case n'est cochée"
    End If
   
    If (CheckBox1.Value = True) Then
        'Recup etatcivil.txt
        i = 0
        Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Etatcivil.txt" For Input As #1
        While Not EOF(1)
            Line Input #1, ContenuLigne
            Result = Split(ContenuLigne, ",")
            If Result(0) = numEtudiant Then
                MsgBox "Etat civil supprimé"
                TextBox1.Value = ""
                TextBox2.Value = ""
                TextBox9.Value = ""
                TextBox3.Value = ""
                TextBox4.Value = ""
            Else
                Tablo(i) = Result(0) & "," & Result(1) & "," & Result(2) & "," & Result(3)
                i = i + 1
            End If
            
        Wend
        Close #1
        Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Etatcivil.txt" For Output As #1
        j = 0
        While j < i
            Print #1, Tablo(j)
            j = j + 1
        Wend
        Close #1
    End If
    
    If (CheckBox2.Value = True) Then
        'Recup notes.txt
        i = 0
        Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Notes.txt" For Input As #1
        While Not EOF(1)
            Line Input #1, ContenuLigne
            Result = Split(ContenuLigne, ",")
            If Result(0) = numEtudiant Then
                MsgBox "Notes supprimées"
                TextBox6.Value = ""
                TextBox5.Value = ""
            Else
                Tablo(i) = Result(0) & "," & Result(1) & "," & Result(2)
                i = i + 1
            End If
        Wend
        Close #1
        Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Notes.txt" For Output As #1
        j = 0
        While j < i
            Print #1, Tablo(j)
            j = j + 1
        Wend
        Close #1
    End If
    
    If (CheckBox3.Value = True) Then
        'Recup adresses.txt
        i = 0
        Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Adresses.txt" For Input As #1
        While Not EOF(1)
            Line Input #1, ContenuLigne
            Result = Split(ContenuLigne, ",")
            If Result(0) = numEtudiant Then
                MsgBox "Adresse supprimée"
                TextBox8.Value = ""
                TextBox7.Value = ""
                
            Else
                Tablo(i) = Result(0) & "," & Result(1) & "," & Result(2)
                i = i + 1
            End If
        Wend
        Close #1
        Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Adresses.txt" For Output As #1
        j = 0
        While j < i
            Print #1, Tablo(j)
            j = j + 1
        Wend
        Close #1
    End If
    
    
    
    
End Sub

Private Sub TextBox5_Change()

End Sub

Private Sub TextBox6_Change()

End Sub

Private Sub TextBox7_Change()

End Sub

Private Sub TextBox8_Change()

End Sub

Private Sub TextBox9_Change()

End Sub

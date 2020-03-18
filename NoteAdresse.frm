VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NoteAdresse 
   Caption         =   "Création et affichage des Notes et Adresses"
   ClientHeight    =   7548
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10392
   OleObjectBlob   =   "NoteAdresse.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NoteAdresse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim num As String, TP As String, DS As String, CP As String, Adresse As String
Dim exist As Boolean, n As Integer, exist2 As Boolean
Dim Result() As String, Result2() As String
Dim identite As String

Private Sub CommandButton1_Click()
    num = TextBox4.Value
    TP = TextBox3.Value
    DS = TextBox2.Value
    exist = False
    exist2 = False
    n = 0
    
    Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Notes.txt" For Input As #2
    
    While Not EOF(2)
        Line Input #2, ContenuLigne
        Result = Split(ContenuLigne, ",")
        If (Result(0) = num) Then
            exist2 = True
        End If
    Wend
    
    Close #2
    
    If (exist2 = False) Then
    Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Etatcivil.txt" For Input As #1
    Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Notes.txt" For Append As #2
    
    While Not EOF(1)
        Line Input #1, ContenuLigne
        Result = Split(ContenuLigne, ",")
        
        If (Result(0) = num) Then
            exist = True
        End If
    Wend
    
    If (exist = True) Then
        If (num <> "" And TP <> "" And DS <> "") Then
            If (IsNumeric(TP) And IsNumeric(DS)) Then
                Print #2, num & "," & TP & "," & DS
                MsgBox "Les notes sont créées"
            Else
                MsgBox "Les notes doivent être numériques"
            End If
        Else
            MsgBox "Il faut remplir tous les champs"
        End If
    Else
        MsgBox "Cet étudiant n'existe pas"
    End If

    Close #1
    Close #2
    
    Else
        MsgBox "Les notes existes déjà pour cet étudiant"
    End If
    
    
    
End Sub

Private Sub CommandButton2_Click()

Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Notes.txt" For Input As #2

ListBox1.Clear

While Not EOF(2)
    Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Etatcivil.txt" For Input As #1
    Line Input #2, ContenuLigne2
    Result = Split(ContenuLigne2, ",")
    While Not EOF(1)
        Line Input #1, ContenuLigne1
        Result2 = Split(ContenuLigne1, ",")
        If (Result2(0) = Result(0)) Then
            identite = Result2(1) & " " & Result2(2)
        End If
    Wend
    ListBox1.AddItem Result(0) & " (" & identite & ") | TP : " & Result(1) & ", DS : " & Result(2)
    identite = ""
    Close #1
    n = n + 1
Wend

If (n = 0) Then
    ListBox1.AddItem "Aucune note"
End If

Close #2

End Sub

Private Sub CommandButton3_Click()
    num = TextBox7.Value
    CP = TextBox6.Value
    Adresse = TextBox5.Value
    exist = False
    exist2 = False
    n = 0
    
    Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Adresses.txt" For Input As #2
    
    While Not EOF(2)
        Line Input #2, ContenuLigne
        Result = Split(ContenuLigne, ",")
        If (Result(0) = num) Then
            exist2 = True
        End If
    Wend
    
    Close #2
    
    If (exist2 = False) Then
        Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Etatcivil.txt" For Input As #1
        Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Adresses.txt" For Append As #2
        
        While Not EOF(1)
            Line Input #1, ContenuLigne
            Result = Split(ContenuLigne, ",")
            
            If (Result(0) = num) Then
                exist = True
            End If
        Wend
        
        If (exist = True) Then
            If (num <> "" And CP <> "" And Adresse <> "") Then
                Print #2, num & "," & CP & "," & Adresse
                MsgBox "L'adresse est créée"
            Else
                MsgBox "Il faut remplir tous les champs"
            End If
        Else
            MsgBox "Cet étudiant n'existe pas"
        End If
    
        Close #1
        Close #2
    Else
        MsgBox "Cet étudiant a déjà une adresse"
    End If
    
End Sub

Private Sub CommandButton4_Click()

Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Adresses.txt" For Input As #2

ListBox2.Clear

While Not EOF(2)
    Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Etatcivil.txt" For Input As #1
    Line Input #2, ContenuLigne2
    Result = Split(ContenuLigne2, ",")
    While Not EOF(1)
        Line Input #1, ContenuLigne1
        Result2 = Split(ContenuLigne1, ",")
        If (Result2(0) = Result(0)) Then
            identite = Result2(1) & " " & Result2(2)
        End If
    Wend
    ListBox2.AddItem Result(0) & " (" & identite & ") | Code Postal : " & Result(1) & ", Adresse : " & Result(2)
    identite = ""
    Close #1
    n = n + 1
Wend

If (n = 0) Then
    ListBox2.AddItem "Aucune adresse"
End If

Close #2

End Sub

Private Sub TextBox6_Change()

End Sub

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Requetes 
   Caption         =   "Tout un tas de requêtes !"
   ClientHeight    =   9048.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8208.001
   OleObjectBlob   =   "Requetes.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Requetes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim addThem(1000) As String
Dim adresses(1000) As String

Private Sub CommandButton1_Click()
ville = TextBox1.Value
i = 0
j = 0

    If (ville <> "") Then
        'Parcours des adresses
        Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Adresses.txt" For Input As #1
        While Not EOF(1)
            Line Input #1, ContenuLigne
            Result = Split(ContenuLigne, ",")
            If (Result(2) = ville) Then
                addThem(i) = Result(0)
                i = i + 1
            End If
        Wend
        
        Close #1
        
        'Affichage
        Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Etatcivil.txt" For Input As #1
    
        ListBox1.Clear
           
        While Not EOF(1)
            j = 0
            Line Input #1, ContenuLigne
            Result = Split(ContenuLigne, ",")
            While j < i
                If (addThem(j) = Result(0)) Then
                    ListBox1.AddItem Result(1) & " " & Result(2)
                End If
                j = j + 1
            Wend
        Wend
    
    Close #1
    
    End If

End Sub


Private Sub CommandButton2_Click()
nom = TextBox2.Value
i = 0
j = 0
k = 0

    If (nom <> "") Then
        'Parcours des noms
        Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Etatcivil.txt" For Input As #1
        While Not EOF(1)
            Line Input #1, ContenuLigne
            Result = Split(ContenuLigne, ",")
            If (Result(1) = nom) Then
                addThem(i) = Result(0)
                i = i + 1
            End If
        Wend
        
        Close #1
        
        'Parcours des adresses
        Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Adresses.txt" For Input As #1
        While Not EOF(1)
            Line Input #1, ContenuLigne
            Result = Split(ContenuLigne, ",")
            k = 0
            While k < i
                If (Result(0) = addThem(k)) Then
                    adresses(k) = Result(2)
                End If
                k = k + 1
            
            Wend
           
        Wend
        
        Close #1
        
        'Affichage
        Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Etatcivil.txt" For Input As #1
    
        ListBox2.Clear
           
        While Not EOF(1)
            j = 0
            Line Input #1, ContenuLigne
            Result = Split(ContenuLigne, ",")
            While j < i
                If (addThem(j) = Result(0)) Then
                    ListBox2.AddItem Result(0) & " - " & Result(1) & " " & Result(2) & " - " & adresses(j)
                End If
                j = j + 1
            Wend
        Wend
    
    Close #1
    
    End If

End Sub

Private Sub CommandButton4_Click()
    nom = TextBox3.Value
    prenom = TextBox4.Value
    num = ""
    exist = False
    
    'Parcours des noms
    If (nom <> "" And prenom <> "") Then
        Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Etatcivil.txt" For Input As #1
        While Not EOF(1)
            Line Input #1, ContenuLigne
            Result = Split(ContenuLigne, ",")
            If (Result(1) = nom And Result(2) = prenom) Then
                num = Result(0)
            End If
        Wend
        Close #1
    Else
        MsgBox "Champs invalides"
    End If
    
    'Parcours des notes et affichage
    If (num <> "") Then
        Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Notes.txt" For Input As #1
        While Not EOF(1)
            Line Input #1, ContenuLigne
            Result = Split(ContenuLigne, ",")
            If (Result(0) = num) Then
                TextBox5.Value = CStr(Round(((CDbl(Result(1)) + 2 * CDbl(Result(2))) / 3), 2))
   
                TextBox5.Value = TextBox5.Value & "/20"
                exist = True
            End If
        Wend
        Close #1
    End If
    
    If (exist = False) Then
        MsgBox "Aucune note pour cet étudiant !"
        TextBox3.Value = ""
        TextBox4.Value = ""
        TextBox5.Value = ""
    End If
    

End Sub

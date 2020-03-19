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
Dim merge(1000) As String


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

Private Sub CommandButton5_Click()
lettre = TextBox7.Value
newLettre = ""
i = 0
c = 0
j = 0

'Parcours des etudiants
Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Etatcivil.txt" For Input As #1

While Not EOF(1)
    Line Input #1, ContenuLigne
    Result = Split(ContenuLigne, ",")
    newLettre = Left(Result(2), 1)
    If (newLettre <> lettre) Then
        addThem(i) = Result(0) & "," & Result(1) & "," & Result(2) & "," & Result(3)
        i = i + 1
    Else
        c = c + 1
    End If
Wend
MsgBox CStr(c) & " étudiants supprimés"

Close #1

Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Etatcivil.txt" For Output As #1

While j < i
    Print #1, addThem(j)
    j = j + 1
Wend

Close #1


End Sub

Private Sub CommandButton6_Click()
'Affichage des etudiants
Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Etatcivil.txt" For Input As #1

ListBox3.Clear


While Not EOF(1)
    Line Input #1, ContenuLigne
    Result = Split(ContenuLigne, ",")
    ListBox3.AddItem "Etudiant " & Result(0) & " : " & Result(1) & " " & Result(2) & ", " & Result(3)
Wend

Close #1
End Sub

Private Sub CommandButton7_Click()
'Variables
i = 0
j = 0
numEtu = ""

'Check Checkboxes
mergeNum = CheckBox1.Value
mergeNom = CheckBox2.Value
mergePrenom = CheckBox5.Value
mergeNatio = CheckBox6.Value
mergeTP = CheckBox3.Value
mergeDS = CheckBox7.Value
mergeCP = CheckBox4.Value
mergeAdresse = CheckBox8.Value

'Check Values
Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Etatcivil.txt" For Input As #1

While Not EOF(1)
    'Etatcivil
    Line Input #1, ContenuLigne
    Result = Split(ContenuLigne, ",")
    numEtu = Result(0)
    If (mergeNum = True) Then
        If (merge(i) = "") Then
            merge(i) = "num:" & Result(0)
        Else
            merge(i) = merge(i) & ", num:" & Result(0)
        End If
    End If
    If (mergeNom) Then
        If (merge(i) = "") Then
            merge(i) = "nom:" & Result(1)
        Else
            merge(i) = merge(i) & ", nom:" & Result(1)
        End If
    End If
    If (mergePrenom) Then
        If (merge(i) = "") Then
            merge(i) = "prenom:" & Result(2)
        Else
            merge(i) = merge(i) & ", prenom:" & Result(2)
        End If
    End If
    If (mergeNatio) Then
        If (merge(i) = "") Then
            merge(i) = "natio:" & Result(3)
        Else
            merge(i) = merge(i) & ", natio:" & Result(3)
        End If
            
    End If
    
    'Notes
    Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Notes.txt" For Input As #2
    While Not EOF(2)
        Line Input #2, ContenuLigne
        Result = Split(ContenuLigne, ",")
        If (Result(0) = numEtu) Then
            If (mergeTP = True) Then
                If (merge(i) = "") Then
                    merge(i) = "TP:" & Result(1)
                Else
                    merge(i) = merge(i) & ", TP:" & Result(1)
                End If
            End If
            If (mergeDS = True) Then
                If (merge(i) = "") Then
                    merge(i) = "DS:" & Result(2)
                Else
                    merge(i) = merge(i) & ", DS:" & Result(2)
                End If
            End If
        End If
    Wend
    Close #2
    
    'Adresses
    Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Adresses.txt" For Input As #3
    While Not EOF(3)
        Line Input #3, ContenuLigne
        Result = Split(ContenuLigne, ",")
        If (Result(0) = numEtu) Then
            If (mergeCP = True) Then
                If (merge(i) = "") Then
                    merge(i) = "CP:" & Result(1)
                Else
                    merge(i) = merge(i) & ", CP:" & Result(1)
                End If
            End If
            If (mergeAdresse = True) Then
                If (merge(i) = "") Then
                    merge(i) = "Adresse:" & Result(2)
                Else
                    merge(i) = merge(i) & ", Adresse:" & Result(2)
                End If
            End If
        End If
    Wend
    Close #3
    
    i = i + 1
    
    
    
Wend

Close #1


'Print in fusion.txt
Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Fusion.txt" For Output As #4
While j < i
    Print #4, merge(j)
    j = j + 1
Wend
Close #4

'Reset array
j = 0
While j < i
    merge(j) = ""
    j = j + 1
Wend

End Sub

Private Sub CommandButton8_Click()
ListBox4.Clear

Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Fusion.txt" For Input As #1
While Not EOF(1)
    Line Input #1, ContenuLigne
    ListBox4.AddItem ContenuLigne
Wend

Close #1
End Sub

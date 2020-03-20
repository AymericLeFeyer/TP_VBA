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
    'Valeurs depuis les TextBox
    num = TextBox4.Value
    TP = TextBox3.Value
    DS = TextBox2.Value
    'Variables utiles apres
    exist = False
    exist2 = False
    n = 0
    'Ouverture du fichier Notes
    Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Notes.txt" For Input As #2
    'On parcours ce fichier
    While Not EOF(2)
        Line Input #2, ContenuLigne
        Result = Split(ContenuLigne, ",")
        'Si le numero existe deja, on s'arretera
        If (Result(0) = num) Then
            exist2 = True
        End If
    Wend
    
    Close #2
    'Si le numero n'existe pas deja
    If (exist2 = False) Then
        'On ouvre les fichiers
        Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Etatcivil.txt" For Input As #1
        Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Notes.txt" For Append As #2
        'On parcours Etatcivil
        While Not EOF(1)
            Line Input #1, ContenuLigne
            Result = Split(ContenuLigne, ",")
            'Si le numero existe, on pourra ajouter des notes
            If (Result(0) = num) Then
                exist = True
            End If
        Wend
        'Le numero existe, on peux ajouter des scores
        If (exist = True) Then
            'On verifie que les champs ne sont pas nulles
            If (num <> "" And TP <> "" And DS <> "") Then
                'On verifie que les champs sont numeriques
                If (IsNumeric(TP) And IsNumeric(DS)) Then
                    'On ajoute le ligne
                    Print #2, num & "," & TP & "," & DS
                    'On affiche un message de confirmation
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

'On ouvre les notes
Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Notes.txt" For Input As #2

'On clear la liste
ListBox1.Clear

'On parcours les notes
While Not EOF(2)
    'On ouvre les etats civils
    Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Etatcivil.txt" For Input As #1
    Line Input #2, ContenuLigne2
    Result = Split(ContenuLigne2, ",")
    'On parcours les etats civils
    While Not EOF(1)
        Line Input #1, ContenuLigne1
        Result2 = Split(ContenuLigne1, ",")
        'On cherche a fusionner les notes et l'etat civil pour un affichage propre
        If (Result2(0) = Result(0)) Then
            'On rentrera forcement dans cette boucle une seule fois, grace a toutes les conditions depuis le debut
            identite = Result2(1) & " " & Result2(2)
        End If
    Wend
    'On affiche la ligne
    ListBox1.AddItem Result(0) & " (" & identite & ") | TP : " & Result(1) & ", DS : " & Result(2)
    identite = ""
    Close #1
    'On incremente un compteur
    n = n + 1
Wend

'Si le compteur est vide, on affiche "Aucune note"
If (n = 0) Then
    ListBox1.AddItem "Aucune note"
End If

Close #2

End Sub

Private Sub CommandButton3_Click()
    'Valeurs depuis les TextBox
    num = TextBox7.Value
    CP = TextBox6.Value
    Adresse = TextBox5.Value
    'Variables utiles apres
    exist = False
    exist2 = False
    n = 0
    'Ouverture des adresses pour verifier qu'une adresse n'existe pas deja
    Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Adresses.txt" For Input As #2
    
    While Not EOF(2)
        Line Input #2, ContenuLigne
        Result = Split(ContenuLigne, ",")
        'On verifie donc que le numero n'a pas deja une adresse
        If (Result(0) = num) Then
            exist2 = True
        End If
    Wend
    
    Close #2
    'Si le numero n'a pas encore d'adresse
    If (exist2 = False) Then
        'On ouvre les fichiers
        Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Etatcivil.txt" For Input As #1
        Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Adresses.txt" For Append As #2
        
        While Not EOF(1)
            Line Input #1, ContenuLigne
            Result = Split(ContenuLigne, ",")
            'On verifie que le numero indique existe
            If (Result(0) = num) Then
                exist = True
            End If
        Wend
        'On verifie qu'aucun champ n'est nul
        If (exist = True) Then
            If (num <> "" And CP <> "" And Adresse <> "") Then
                'On ajoute la ligne
                Print #2, num & "," & CP & "," & Adresse
                'On affiche un message
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
'On vide la liste
ListBox2.Clear
'On parcours les adresses
While Not EOF(2)
    Open "C:\Users\Aymeric\Documents\GitHub\TP_VBA" & "\Etatcivil.txt" For Input As #1
    Line Input #2, ContenuLigne2
    Result = Split(ContenuLigne2, ",")
    'On fais la liaison avec l'identite de la personne pour un affichage propre, on dois rentrer une fois dans la condition If plus loin
    While Not EOF(1)
        Line Input #1, ContenuLigne1
        Result2 = Split(ContenuLigne1, ",")
        If (Result2(0) = Result(0)) Then
            identite = Result2(1) & " " & Result2(2)
        End If
    Wend
    'On ajoute la ligne a la liste
    ListBox2.AddItem Result(0) & " (" & identite & ") | Code Postal : " & Result(1) & ", Adresse : " & Result(2)
    identite = ""
    Close #1
    n = n + 1
Wend

'On affiche ce message si aucune adresse existe
If (n = 0) Then
    ListBox2.AddItem "Aucune adresse"
End If

Close #2

End Sub

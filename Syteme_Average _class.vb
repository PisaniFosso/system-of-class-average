'******************************************************
'Auteur :  Pisani Fosso A00187566                    **
'******************************************************
'Nom du programme: calculatrice de moyenne de classe **
'******************************************************
'But: comprendre comment utiliser des «ListBox»,     **
'     les boucles et les fonctions.                  **
'******************************************************
Public Class Form1

    'cliquez sur Quitter pour fermer le programme
    Private Sub Button_Exit_Click(sender As Object, e As EventArgs) Handles Button_Exit.Click

        Me.Close()
    End Sub

    'Cliquer sur Ajout pour:
    '- prendre le nom de l'étudiant entré et l'ajouter dans ListBox_Students qui est la liste de tous les élèves
    '- prendre la note choisi dans le comboBox des notes et l'ajouter dans ListBox_Average qui est la moyenne de tous les élèves
    '- effacer le nom entée précédement
    'si l'étudiant n'a pas de nom ou de note un message d'erreur s'affichera 
    Private Sub Button_Add_Click(sender As Object, e As EventArgs) Handles Button_Add.Click

        If TextBox_StudentName.Text <> "" And ComboBox_Averages.Text <> "" Then
            ListBox_Students.Items.Add(TextBox_StudentName.Text)
            ListBox_Average.Items.Add(ComboBox_Averages.Text)
            TextBox_StudentName.Text = ""
        Else
            MessageBox.Show("Remplir le Nom de l'étudiant ou la note", "Erreur")
        End If

    End Sub

    'Lorsqu'un etudiant est sélectionner, sa note est automatiquement selectionner aussi
    'un message indiquant son nom et sa note apparait dans Label_Message
    Private Sub ListBox_Students_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox_Students.SelectedIndexChanged

        ListBox_Average.SelectedIndex = ListBox_Students.SelectedIndex
        Select Case ListBox_Average.SelectedItem
            Case "0-59"     'si la note est entre 0-59 la note de l'étudiant est E
                Label_Message.Text = "La note à " & ListBox_Students.SelectedItem & " est E."
            Case "60-69"    'si la note est entre 60-69 la note de l'étudiant est D
                Label_Message.Text = "La note à " & ListBox_Students.SelectedItem & " est D."
            Case "70-79"    'si la note est entre 70-79 la note de l'étudiant est C
                Label_Message.Text = "La note à " & ListBox_Students.SelectedItem & " est C."
            Case "80-89"    'si la note est entre 80-89 la note de l'étudiant est B
                Label_Message.Text = "La note à " & ListBox_Students.SelectedItem & " est B."
            Case "90-100"   'si la note est entre 90-100 la note de l'étudiant est A
                Label_Message.Text = "La note à " & ListBox_Students.SelectedItem & " est A."
        End Select
    End Sub

    'Lorsqu'on clique sur le boutton Supprimer,
    'si un étudiant est selectionné, son nom et sa moyenne sont retirés de nos listes;
    'sinon afficher un message d'erreur
    Private Sub Button_Delete_Click(sender As Object, e As EventArgs) Handles Button_Delete.Click

        If ListBox_Students.SelectedItem <> "" Then
            ListBox_Average.Items.RemoveAt(ListBox_Students.SelectedIndex)
            ListBox_Students.Items.Remove(ListBox_Students.SelectedItem)
            Label_Message.Text = ""
        Else
            MessageBox.Show("Selectionner l'étudiant à retirer", "Erreur")
        End If
    End Sub

    'Lorsqu'on clique sur le boutton Vider,
    '- Les listes de noms et celle de moyenne sont vider totalement;
    '- La moyenne et le nombre d'etudiants deviennent vide
    Private Sub Button_Clear_Click(sender As Object, e As EventArgs) Handles Button_Clear.Click

        ListBox_Students.Items.Clear()
        ListBox_Average.Items.Clear()
        Label_Average.Text = ""
        Label_NumberOfStudents.Text = ""
    End Sub

    'Lorsqu'on clique sur le boutton Calcul,
    '- La fonction Calcul_Average est appelé pour calculer la moyenne,
    '- La moyenne et le nombre d'étudiants de la classse sont mis a jour respectivement
    '  dans Label_Average et Label_NumberOfStudents
    Private Sub Button3_Calcul_Click(sender As Object, e As EventArgs) Handles Button3_Calcul.Click

        Label_NumberOfStudents.Text = ListBox_Students.Items.Count
        Dim Average As Double
        Calcul_Average(Average)
        ListBox_Students.SelectedIndex = ListBox_Average.SelectedIndex
        Select Case Average
            Case 0 To 59    'si Average est entre 0-59 la Moyenne de la classe est E
                Label_Average.Text = "0-59"
            Case 60 To 69   'si Average est entre 60-69 la Moyenne de la classe est C
                Label_Average.Text = "60-69"
            Case 70 To 79   'si Average est entre 70-79 la Moyenne de la classe est D
                Label_Average.Text = "70-79"
            Case 80 To 89   'si Average est entre 80-89 la Moyenne de la classe est B
                Label_Average.Text = "80-89"
            Case 90 To 100  'si Average est entre 90-100 la Moyenne de la classe est A
                Label_Average.Text = "90-100"
        End Select
    End Sub

    'Calcul_Average ne prend rien en paramettre
    'retourne un Double
    'La fonction ci-dessous prend la moyenne de chaque étudiants dans ListBox_Average,
    'Split le text pour extraire le nombre après le tiret qui sera la moyenne de l'étudiant,
    'les moyennes extraites de chaque seront par la suite additioné 
    'puis divisé par le nombre total d'étudiants
    Private Sub Calcul_Average(ByRef todoAverage As Double)

        Dim txAverage As String

        'du premier index qui est 0 au dernier index qui est le nombre
        'retirer les moyennes et les additioner
        For index = 0 To ListBox_Students.Items.Count - 1
            ListBox_Average.SelectedIndex = index
            txAverage = ListBox_Average.SelectedItem
            Dim Average As String() = txAverage.Split(New Char() {"-"c})
            Dim Averagedl As Double = Average(1)
            todoAverage += Averagedl    '<=> todoAverage = todoAverage + Averagedl
        Next index      'Incrémenter

        todoAverage = (todoAverage / ListBox_Students.Items.Count)

    End Sub

End Class

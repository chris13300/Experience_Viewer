Public Class frmOptions

    Public tabDef() As String
    Public tabMin() As String
    Public tabMax() As String

    Private Sub frmOptions_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim tabChaine() As String, i As Integer, tabTmp() As String, j As Integer

        tabChaine = Split(optionsUCI(entree, sortie), vbCrLf)
        ReDim tabDef(UBound(tabChaine))
        ReDim tabMin(UBound(tabChaine))
        ReDim tabMax(UBound(tabChaine))
        lstNoms.Items.Clear()
        lstValeurs.Items.Clear()
        lstLog.Items.Clear()
        For i = 0 To UBound(tabChaine)
            If tabChaine(i) <> "" Then
                tabTmp = Split(tabChaine(i), "|")
                If tabTmp.Length = 2 Then
                    tabDef(i) = tabTmp(1)
                    tabMin(i) = ""
                    tabMax(i) = ""
                    lstNoms.Items.Add(tabTmp(0))
                Else
                    tabDef(i) = tabTmp(1)
                    tabMin(i) = tabTmp(2)
                    tabMax(i) = tabTmp(3)
                    lstNoms.Items.Add(tabTmp(0) & " (min. " & tabMin(i) & ", max. " & tabMax(i) & ")")
                End If
                If LCase(tabTmp(0)) = "experience file" Then
                    tabTmp(1) = frmPrincipale.lblExperience.Text
                End If
                lstValeurs.Items.Add(tabTmp(1))
            End If
        Next

        If My.Computer.FileSystem.FileExists(Replace(moteur_court, ".exe", ".ini")) Then
            tabChaine = Split(My.Computer.FileSystem.ReadAllText(Replace(moteur_court, ".exe", ".ini")), vbCrLf)
            For i = 0 To UBound(tabChaine)
                If InStr(tabChaine(i), " = ") > 0 Then
                    tabTmp = Split(tabChaine(i), " = ")
                    For j = 0 To lstNoms.Items.Count - 1
                        If InStr(lstNoms.Items(j), " (min") > 0 Then
                            If gauche(lstNoms.Items(j), lstNoms.Items(j).indexof(" (min")) = tabTmp(0) Then
                                lstValeurs.Items(j) = tabTmp(1)
                            End If
                        Else
                            If lstNoms.Items(j) = tabTmp(0) Then
                                lstValeurs.Items(j) = tabTmp(1)
                            End If
                        End If
                        If LCase(lstNoms.Items(j)) = "experience file" Then
                            lstValeurs.Items(j) = frmPrincipale.lblExperience.Text
                        End If
                    Next
                End If
            Next
        End If
    End Sub

    Private Sub lstValeurs_Click(sender As Object, e As EventArgs) Handles lstValeurs.Click
        lstNoms.SetSelected(lstValeurs.SelectedIndex, True)
    End Sub

    Private Sub lstValeurs_DoubleClick(sender As Object, e As EventArgs) Handles lstValeurs.DoubleClick
        Dim reponse As String
        If lstValeurs.SelectedItem = "false" Then
            lstValeurs.Items(lstValeurs.SelectedIndex) = "true"
        ElseIf lstValeurs.SelectedItem = "true" Then
            lstValeurs.Items(lstValeurs.SelectedIndex) = "false"
        Else
            reponse = InputBox("New value ?", lstNoms.Items(lstNoms.SelectedIndex), lstValeurs.SelectedItem)
            If reponse <> "" And reponse <> lstValeurs.SelectedItem Then
                If IsNumeric(reponse) Then
                    If tabMin(lstValeurs.SelectedIndex) <= Val(reponse) And Val(reponse) <= tabMax(lstValeurs.SelectedIndex) Then
                        lstValeurs.Items(lstValeurs.SelectedIndex) = reponse
                    Else
                        MsgBox("min. " & tabMin(lstValeurs.SelectedIndex) & ", max. " & tabMax(lstValeurs.SelectedIndex))
                    End If
                Else
                    lstValeurs.Items(lstValeurs.SelectedIndex) = reponse
                End If
            End If
        End If
    End Sub

    Private Sub cmdAppliquer_Click(sender As Object, e As EventArgs) Handles cmdAppliquer.Click
        Dim i As Integer, modif As String, chaine As String

        modif = ""
        For i = 0 To lstValeurs.Items.Count - 1
            If lstNoms.Items(i) <> "" And lstValeurs.Items(i) <> tabDef(i) Then
                If InStr(lstNoms.Items(i), " (min") > 0 Then
                    chaine = "setoption name " & gauche(lstNoms.Items(i), lstNoms.Items(i).indexof(" (min")) & " value " & lstValeurs.Items(i)
                    modif = modif & gauche(lstNoms.Items(i), lstNoms.Items(i).indexof(" (min")) & " = " & lstValeurs.Items(i) & vbCrLf
                Else
                    chaine = "setoption name " & lstNoms.Items(i) & " value " & lstValeurs.Items(i)
                    modif = modif & lstNoms.Items(i) & " = " & lstValeurs.Items(i) & vbCrLf
                End If
                entree.WriteLine(chaine)
                tabDef(i) = lstValeurs.Items(i)
                lstLog.Items.Add(chaine)
                lstLog.SetSelected(lstLog.Items.Count - 1, True)
            End If
        Next
        If modif <> "" Then
            My.Computer.FileSystem.WriteAllText(Replace(moteur_court, ".exe", ".ini"), modif, False)
            entree.WriteLine("isready")
            chaine = ""
            While InStr(chaine, "readyok") = 0
                chaine = sortie.ReadLine
                If chaine <> "readyok" Then
                    lstLog.Items.Add(chaine)
                    lstLog.SetSelected(lstLog.Items.Count - 1, True)
                End If
                Threading.Thread.Sleep(1)
            End While
        End If

        chaine = InputBox("Min. depth ?", "analysis", profMinAnalyse)
        If IsNumeric(chaine) Then
            profMinAnalyse = CInt(chaine)
            If profMaxAnalyse < profMinAnalyse Then
                profMaxAnalyse = profMinAnalyse
            End If
        End If

        chaine = InputBox("Max. depth ?", "analysis", profMaxAnalyse)
        If IsNumeric(chaine) Then
            profMaxAnalyse = CInt(chaine)
        End If

        chaine = InputBox("Mode ?" & vbCrLf & "1 : all legal moves" & vbCrLf & "2 : only played moves" & vbCrLf & "3 : only moves with positive score" & vbCrLf & "4 : search the bestmove", "analysis", modeAnalyse)
        If IsNumeric(chaine) Then
            modeAnalyse = CInt(chaine)
        End If

        analyse = True
        Me.Hide()
    End Sub

End Class
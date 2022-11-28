Imports System.Threading

Public Class frmPrincipale
    Public tabEPD() As String
    Public indexEPD As Integer
    Public tabPGN() As String
    Public tabStartpos() As String
    Public indexPGN As Integer
    Public tabCoups() As String
    Public indexCoup As Integer
    Public fichierINI As String
    Public fichierUCI As String
    Public tabExperience() As Byte
    Public chargement As Boolean
    Public tacheAnalyse As Thread
    Public coup1 As String
    Public coup2 As String
    Public listeExperience As String

    Private Sub frmPrincipale_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        chargement = False
        coup1 = ""
        coup2 = ""
        profMinAnalyse = 16
        profMaxAnalyse = 32
        modeAnalyse = 1
    End Sub

    Private Sub frmPrincipale_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        Dim chaine As String, tabChaine() As String, tabTmp() As String, i As Integer

        Me.Tag = Me.Text

        If Not chargement Then
            chargement = True

            ReDim tabEPD(0)
            indexEPD = 0
            tabEPD(indexEPD) = "rnbqkbnr/pppppppp/8/8/8/8/PPPPPPPP/RNBQKBNR w KQkq - 0 1"
            indexPGN = -1
            lblInfo.Text = tabEPD(indexEPD)
            fichierUCI = ""
            moteur_court = ""

            pbEchiquier.BackgroundImage = Image.FromFile("BMP\echiquier.bmp")
            pbMateriel.BackColor = Color.FromName("control")

            fichierINI = My.Computer.Name & ".ini"
            If My.Computer.FileSystem.FileExists(fichierINI) Then
                chaine = My.Computer.FileSystem.ReadAllText(fichierINI)
                If chaine <> "" And InStr(chaine, vbCrLf) > 0 Then
                    tabChaine = Split(chaine, vbCrLf)
                    For i = 0 To UBound(tabChaine)
                        If tabChaine(i) <> "" And InStr(tabChaine(i), " = ") > 0 Then
                            tabTmp = Split(tabChaine(i), " = ")
                            If tabTmp(0) <> "" And tabTmp(1) <> "" Then
                                If InStr(tabTmp(1), "//") > 0 Then
                                    tabTmp(1) = Trim(gauche(tabTmp(1), tabTmp(1).IndexOf("//") - 1))
                                End If
                                Select Case tabTmp(0)
                                    Case "moteurEXP"
                                        If My.Computer.FileSystem.FileExists(tabTmp(1)) Then
                                            lblEman.Text = tabTmp(1)
                                        End If

                                    Case "fichierEXP"
                                        If My.Computer.FileSystem.FileExists(tabTmp(1)) Then
                                            lblExperience.Text = tabTmp(1)
                                        End If

                                    Case "profMinAnalyse"
                                        profMinAnalyse = CInt(tabTmp(1))

                                    Case "profMaxAnalyse"
                                        profMaxAnalyse = CInt(tabTmp(1))

                                    Case "modeAnalyse"
                                        modeAnalyse = CInt(tabTmp(1))

                                    Case Else

                                End Select
                            End If
                        End If
                        Application.DoEvents()
                    Next
                End If
            End If
            My.Computer.FileSystem.WriteAllText(fichierINI, "moteurEXP = " & lblEman.Text & vbCrLf _
                                                          & "fichierEXP = " & lblExperience.Text & vbCrLf _
                                                          & "profMinAnalyse = " & profMinAnalyse & vbCrLf _
                                                          & "profMaxAnalyse = " & profMaxAnalyse & vbCrLf _
                                                          & "modeAnalyse = " & modeAnalyse, False)

            cmdMoteur.Enabled = False
            If My.Computer.FileSystem.FileExists(lblExperience.Text) Then
                Me.Text = "Loading " & nomFichier(lblExperience.Text) & "..."
                tabExperience = My.Computer.FileSystem.ReadAllBytes(lblExperience.Text)

                cmdMoteur.Enabled = True
                cmdRecherche.Enabled = False
                cmdAnalyse.Enabled = False
                If My.Computer.FileSystem.FileExists(lblEman.Text) Then
                    moteur_court = nomFichier(lblEman.Text)
                    Me.Text = "Loading " & moteur_court & "..."
                    If InStr(moteur_court, "aurora", CompareMethod.Text) > 0 Then
                        lblMoteur.Text = "Aurora EXE :"
                    ElseIf InStr(moteur_court, "sugar", CompareMethod.Text) > 0 Then
                        lblMoteur.Text = "Sugar EXE :"
                    ElseIf InStr(moteur_court, "hypnos", CompareMethod.Text) > 0 Then
                        lblMoteur.Text = "Hypnos EXE :"
                    ElseIf InStr(moteur_court, "stockfishmz", CompareMethod.Text) > 0 Then
                        lblMoteur.Text = "SF MZ EXE :"
                    End If
                    chargerMoteur(lblEman.Text, lblExperience.Text)
                    lblInfo.Text = entete

                    cmdRecherche.Enabled = True
                    cmdAnalyse.Enabled = True
                End If

                If indexEPD >= 0 And moteur_court <> "" Then
                    If tabEPD(indexEPD) <> "" Then
                        afficherEXP(tabEPD(indexEPD))
                    End If
                End If
            End If
        End If

        Me.Text = Me.Tag
    End Sub

    Private Sub frmPrincipale_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If lblEman.Text <> "" Then
            dechargerMoteur()
            If My.Computer.FileSystem.FileExists(fichierUCI) Then
                My.Computer.FileSystem.DeleteFile(fichierUCI)
            End If
        End If
    End Sub

    Private Sub pbEchiquier_Paint(sender As Object, e As PaintEventArgs) Handles pbEchiquier.Paint
        If indexEPD >= 0 And moteur_court <> "" Then
            Try
                echiquier(tabEPD(indexEPD), e, coup1, coup2)
                pbMateriel.Refresh()
            Catch ex As Exception

            End Try
        End If
    End Sub

    Private Sub afficherEXP(positionEPD As String)
        Dim tabChaine() As String, tabTmp() As String, i As Integer

        listeExperience = expListe(positionEPD)
        If listeExperience = "" Then
            For i = 1 To 20
                Me.Controls("lblRang" & i).Text = Format(i, "00")
                Me.Controls("lblCoup" & i).Text = ""
                Me.Controls("lblScore" & i).Text = ""
                Me.Controls("lblScore" & i).ForeColor = Color.Black
                Me.Controls("lblVisite" & i).Text = ""
                Me.Controls("lblQualite" & i).Text = ""
                Application.DoEvents()
            Next
            Exit Sub
        End If
        tabChaine = Split(listeExperience, vbCrLf)
        For i = 1 To 20
            Me.Controls("lblRang" & i).Text = Format(i, "00")
            If i < tabChaine.Length Then
                tabTmp = Split(Replace(tabChaine(i - 1), ":", ","), ",")

                Me.Controls("lblCoup" & i).Text = Trim(tabTmp(1))

                If InStr(tabTmp(5), "cp") > 0 Then
                    'white pov
                    If InStr(positionEPD, " w ") > 0 Then
                        tabTmp(5) = Format(Val(Replace(tabTmp(5), "cp ", "")) / 100, "0.00")
                    ElseIf InStr(positionEPD, " b ") > 0 Then
                        tabTmp(5) = Format(-Val(Replace(tabTmp(5), "cp ", "")) / 100, "0.00")
                    End If
                    If CSng(tabTmp(5)) > 0 Then
                        tabTmp(5) = "+" & tabTmp(5)
                    End If
                ElseIf InStr(tabTmp(5), "mate") > 0 Then
                    'white pov
                    If InStr(positionEPD, " w ") > 0 Then
                        tabTmp(5) = Replace(tabTmp(5), "mate ", "")
                    ElseIf InStr(positionEPD, " b ") > 0 Then
                        tabTmp(5) = Format(-CInt(Replace(tabTmp(5), "mate ", "")))
                    End If
                    If CInt(tabTmp(5)) > 0 Then
                        tabTmp(5) = "+M" & Trim(tabTmp(5))
                    ElseIf CInt(tabTmp(5)) < 0 Then
                        tabTmp(5) = "-M" & Trim(Replace(tabTmp(5), "-", ""))
                    End If
                End If
                If InStr(tabTmp(5), "+") > 0 Then
                    If InStr(positionEPD, " w ", CompareMethod.Text) > 0 Then
                        Me.Controls("lblScore" & i).ForeColor = Color.Blue
                    ElseIf InStr(positionEPD, " b ", CompareMethod.Text) > 0 Then
                        Me.Controls("lblScore" & i).ForeColor = Color.Red
                    End If
                ElseIf InStr(tabTmp(5), "-") > 0 Then
                    If InStr(positionEPD, " w ", CompareMethod.Text) > 0 Then
                        Me.Controls("lblScore" & i).ForeColor = Color.Red
                    ElseIf InStr(positionEPD, " b ", CompareMethod.Text) > 0 Then
                        Me.Controls("lblScore" & i).ForeColor = Color.Blue
                    End If
                Else
                    Me.Controls("lblScore" & i).ForeColor = Color.Black
                End If
                Me.Controls("lblScore" & i).Text = "{" & tabTmp(5) & "/" & Trim(tabTmp(3)) & "}"

                Me.Controls("lblVisite" & i).Text = Format(Val(tabTmp(7)), "##0x")

                Me.Controls("lblQualite" & i).Text = "Q" & Trim(tabTmp(9))
            Else
                Me.Controls("lblCoup" & i).Text = ""
                Me.Controls("lblScore" & i).Text = ""
                Me.Controls("lblScore" & i).ForeColor = Color.Black
                Me.Controls("lblVisite" & i).Text = ""
                Me.Controls("lblQualite" & i).Text = ""
            End If
            Application.DoEvents()
        Next
    End Sub

    Private Sub pbMateriel_Paint(sender As Object, e As PaintEventArgs) Handles pbMateriel.Paint
        If indexEPD >= 0 And moteur_court <> "" Then
            echiquier_differences(tabEPD(indexEPD), e)
        End If
    End Sub

    Private Sub cmdFichier_Click(sender As Object, e As EventArgs) Handles cmdFichier.Click
        Me.Tag = Me.Text

        dlgImport.FileName = ""
        dlgImport.InitialDirectory = Environment.CurrentDirectory
        dlgImport.InitialDirectory = "D:\JEUX\ARENA CHESS 3.5.1\Engines"
        If My.Computer.Name = "WORKSTATION" Or My.Computer.Name = "BUREAU" Or My.Computer.Name = "TOUR-COURTOISIE" Then
            dlgImport.InitialDirectory = "E:\JEUX\ARENA CHESS 3.5.1\Engines"
        End If
        dlgImport.Filter = "Fichier EXP (*.exp)|*.exp"
        dlgImport.ShowDialog()
        If dlgImport.FileName = "" Then
            Exit Sub
        End If

        lblExperience.Text = dlgImport.FileName
        My.Computer.FileSystem.WriteAllText(fichierINI, "moteurEXP = " & lblEman.Text & vbCrLf _
                                                      & "fichierEXP = " & lblExperience.Text & vbCrLf _
                                                      & "profMinAnalyse = " & profMinAnalyse & vbCrLf _
                                                      & "profMaxAnalyse = " & profMaxAnalyse & vbCrLf _
                                                      & "modeAnalyse = " & modeAnalyse, False)

        cmdMoteur.Enabled = False
        If My.Computer.FileSystem.FileExists(lblExperience.Text) Then
            Me.Text = "Loading " & nomFichier(lblExperience.Text) & "..."
            tabExperience = My.Computer.FileSystem.ReadAllBytes(lblExperience.Text)
            If moteur_court <> "" Then
                Me.Text = "Loading " & moteur_court & "..."
                changerExperience(lblExperience.Text)
                lblInfo.Text = entete
                cmdDefrag.Visible = False
                If entete <> "" Then
                    If (InStr(moteur_court, "eman 6", CompareMethod.Text) > 0 And InStr(entete, "Collisions: 0", CompareMethod.Text) = 0) _
                    Or ((InStr(moteur_court, "eman 7", CompareMethod.Text) > 0 Or InStr(moteur_court, "eman 8", CompareMethod.Text) > 0) And InStr(entete, "Duplicate moves: 0", CompareMethod.Text) = 0 _
                    Or ((InStr(moteur_court, "hypnos", CompareMethod.Text) > 0 Or InStr(moteur_court, "stockfishmz", CompareMethod.Text) > 0 Or InStr(moteur_court, "aurora", CompareMethod.Text) > 0) And InStr(entete, "Duplicate moves: 0", CompareMethod.Text) = 0)) Then
                        cmdDefrag.Visible = True
                    End If
                End If
                pbEchiquier.Refresh()
                If indexEPD >= 0 And moteur_court <> "" Then
                    If tabEPD(indexEPD) <> "" Then
                        afficherEXP(tabEPD(indexEPD))
                    End If
                End If
            Else
                cmdRecherche.Enabled = False
                cmdAnalyse.Enabled = False
                If My.Computer.FileSystem.FileExists(lblEman.Text) Then
                    moteur_court = nomFichier(lblEman.Text)
                    Me.Text = "Loading " & moteur_court & "..."
                    If InStr(moteur_court, "aurora", CompareMethod.Text) > 0 Then
                        lblMoteur.Text = "Aurora EXE :"
                    ElseIf InStr(moteur_court, "sugar", CompareMethod.Text) > 0 Then
                        lblMoteur.Text = "Sugar EXE :"
                    ElseIf InStr(moteur_court, "hypnos", CompareMethod.Text) > 0 Then
                        lblMoteur.Text = "Hypnos EXE :"
                    ElseIf InStr(moteur_court, "stockfishmz", CompareMethod.Text) > 0 Then
                        lblMoteur.Text = "SF MZ EXE :"
                    End If
                    chargerMoteur(lblEman.Text, lblExperience.Text)
                    lblInfo.Text = entete
                    cmdDefrag.Visible = False
                    If entete <> "" Then
                        If (InStr(moteur_court, "eman 6", CompareMethod.Text) > 0 And InStr(entete, "Collisions: 0", CompareMethod.Text) = 0) _
                        Or ((InStr(moteur_court, "eman 7", CompareMethod.Text) > 0 Or InStr(moteur_court, "eman 8", CompareMethod.Text) > 0) And InStr(entete, "Duplicate moves: 0", CompareMethod.Text) = 0 _
                        Or ((InStr(moteur_court, "hypnos", CompareMethod.Text) > 0 Or InStr(moteur_court, "stockfishmz", CompareMethod.Text) > 0 Or InStr(moteur_court, "aurora", CompareMethod.Text) > 0) And InStr(entete, "Duplicate moves: 0", CompareMethod.Text) = 0)) Then
                            cmdDefrag.Visible = True
                        End If
                    End If
                    cmdRecherche.Enabled = True
                    cmdAnalyse.Enabled = True
                    pbEchiquier.Refresh()
                    If indexEPD >= 0 And moteur_court <> "" Then
                        If tabEPD(indexEPD) <> "" Then
                            afficherEXP(tabEPD(indexEPD))
                        End If
                    End If
                End If
            End If
            cmdMoteur.Enabled = True
        End If

        Me.Text = Me.Tag
    End Sub

    Private Sub cmdMoteur_Click(sender As Object, e As EventArgs) Handles cmdMoteur.Click
        Me.Tag = Me.Text

        dlgImport.FileName = ""
        dlgImport.InitialDirectory = Environment.CurrentDirectory
        dlgImport.InitialDirectory = "D:\JEUX\ARENA CHESS 3.5.1\Engines"
        If My.Computer.Name = "WORKSTATION" Or My.Computer.Name = "BUREAU" Or My.Computer.Name = "TOUR-COURTOISIE" Then
            dlgImport.InitialDirectory = "E:\JEUX\ARENA CHESS 3.5.1\Engines"
        End If
        dlgImport.Filter = "Moteur EXE (*.exe)|*.exe"
        dlgImport.ShowDialog()
        If dlgImport.FileName = "" Then
            Exit Sub
        End If

        lblEman.Text = dlgImport.FileName
        My.Computer.FileSystem.WriteAllText(fichierINI, "moteurEXP = " & lblEman.Text & vbCrLf _
                                                      & "fichierEXP = " & lblExperience.Text & vbCrLf _
                                                      & "profMinAnalyse = " & profMinAnalyse & vbCrLf _
                                                      & "profMaxAnalyse = " & profMaxAnalyse & vbCrLf _
                                                      & "modeAnalyse = " & modeAnalyse, False)

        cmdRecherche.Enabled = False
        cmdAnalyse.Enabled = False
        If My.Computer.FileSystem.FileExists(lblEman.Text) Then
            moteur_court = nomFichier(lblEman.Text)
            Me.Text = "Loading " & moteur_court & "..."
            If InStr(moteur_court, "aurora", CompareMethod.Text) > 0 Then
                lblMoteur.Text = "Aurora EXE :"
            ElseIf InStr(moteur_court, "sugar", CompareMethod.Text) > 0 Then
                lblMoteur.Text = "SugaR EXE :"
            ElseIf InStr(moteur_court, "hypnos", CompareMethod.Text) > 0 Then
                lblMoteur.Text = "HypnoS EXE :"
            ElseIf InStr(moteur_court, "stockfishmz", CompareMethod.Text) > 0 Then
                lblMoteur.Text = "SF MZ EXE :"
            End If
            chargerMoteur(lblEman.Text, lblExperience.Text)
            lblInfo.Text = entete

            cmdRecherche.Enabled = True
            cmdAnalyse.Enabled = True
        End If

        Me.Text = Me.Tag
    End Sub

    Private Sub cmdRecherche_Click(sender As Object, e As EventArgs) Handles cmdRecherche.Click
        Dim chaineUCI As String, chaineStartpos As String, blanc As String, noir As String, resultat As String, eco As String, startpos As String

        If txtRecherche.Text <> "" Then
            If My.Computer.FileSystem.FileExists(fichierUCI) Then
                My.Computer.FileSystem.DeleteFile(fichierUCI)
                lstParties.Items.Clear()
            End If
        End If

        dlgImport.FileName = ""
        dlgImport.InitialDirectory = Environment.CurrentDirectory
        dlgImport.Filter = "Fichier PGN (*.pgn)|*.pgn|Fichier EPD (*.epd)|*.epd"
        dlgImport.ShowDialog()
        If dlgImport.FileName = "" Then
            Exit Sub
        End If

        txtRecherche.Text = dlgImport.FileName
        If extensionFichier(txtRecherche.Text) = ".pgn" Then
            cmdLignePrecedente.Text = "<<"
            cmdLignePrecedente.Visible = True
            cmdCoupPrecedent.Visible = True
            cmdCoupSuivant.Visible = True
            cmdLigneSuivante.Text = ">>"
            cmdLigneSuivante.Visible = True
            lblRecherche.Text = "Mode PGN :"

            fichierUCI = Replace(txtRecherche.Text, ".pgn", "_uci.pgn")
            If My.Computer.FileSystem.FileExists(fichierUCI) Then
                My.Computer.FileSystem.DeleteFile(fichierUCI)
            End If
            pgnUCI("pgn-extract.exe", txtRecherche.Text, "_uci")
            tabPGN = Split(My.Computer.FileSystem.ReadAllText(fichierUCI), vbCrLf)
            chaineUCI = ""
            chaineStartpos = ""
            blanc = ""
            noir = ""
            resultat = ""
            eco = "???"
            startpos = "rnbqkbnr/pppppppp/8/8/8/8/PPPPPPPP/RNBQKBNR w KQkq - 0 1"
            Me.Width = 1280
            lstParties.Visible = True
            Me.CenterToScreen()
            For i = 0 To UBound(tabPGN)
                If tabPGN(i) <> "" Then
                    If InStr(tabPGN(i), "[White ") > 0 Then
                        blanc = Replace(Replace(tabPGN(i), "[White """, ""), """]", "")
                    ElseIf InStr(tabPGN(i), "[Black ") > 0 Then
                        noir = Replace(Replace(tabPGN(i), "[Black """, ""), """]", "")
                    ElseIf InStr(tabPGN(i), "[Result ") > 0 Then
                        resultat = Replace(Replace(tabPGN(i), "[Result """, ""), """]", "")
                    ElseIf InStr(tabPGN(i), "[ECO ") > 0 Then
                        eco = Replace(Replace(tabPGN(i), "[ECO """, ""), """]", "")
                    ElseIf InStr(tabPGN(i), "[FEN ") > 0 Then
                        startpos = Replace(Replace(tabPGN(i), "[FEN """, ""), """]", "")
                    ElseIf InStr(tabPGN(i), "[") = 0 And InStr(tabPGN(i), "]") = 0 Then
                        tabPGN(i) = Trim(Replace(tabPGN(i), "*", ""))
                        tabPGN(i) = Trim(Replace(tabPGN(i), "1-0", ""))
                        tabPGN(i) = Trim(Replace(tabPGN(i), "0-1", ""))
                        tabPGN(i) = Trim(Replace(tabPGN(i), "1/2-1/2", ""))
                        lstParties.Items.Add(Format(lstParties.Items.Count + 1, "000") & " (" & eco & ", " & Format(tabPGN(i).Split(" ").Length, "000 plies") & ") : " & blanc & " - " & noir & " (" & Replace(resultat, "1/2-1/2", "=-=") & ") " & tabPGN(i))
                        chaineUCI = chaineUCI & tabPGN(i) & vbCrLf
                        chaineStartpos = chaineStartpos & startpos & vbCrLf
                        blanc = ""
                        noir = ""
                        resultat = ""
                        eco = "???"
                        startpos = "rnbqkbnr/pppppppp/8/8/8/8/PPPPPPPP/RNBQKBNR w KQkq - 0 1"
                    End If
                End If
                Application.DoEvents()
            Next
            tabPGN = Split(chaineUCI, vbCrLf)
            tabStartpos = Split(chaineStartpos, vbCrLf)
            indexPGN = -1
        ElseIf extensionFichier(txtRecherche.Text) = ".epd" Then
            cmdLignePrecedente.Text = "<"
            cmdLignePrecedente.Visible = True
            cmdCoupPrecedent.Visible = False
            cmdCoupSuivant.Visible = False
            cmdLigneSuivante.Text = ">"
            cmdLigneSuivante.Visible = True
            lblRecherche.Text = "Mode EPD :"
            tabEPD = Split(My.Computer.FileSystem.ReadAllText(txtRecherche.Text), vbCrLf)
            Me.Width = 1280
            lstParties.Visible = True
            Me.CenterToScreen()
            For i = 0 To UBound(tabEPD)
                If tabEPD(i) <> "" Then
                    lstParties.Items.Add(Format(lstParties.Items.Count + 1, "000") & " : " & tabEPD(i))
                End If
                Application.DoEvents()
            Next
            indexEPD = -1
        End If
        cmdLigneSuivante_Click(sender, e)
    End Sub

    Private Sub cmdLignePrecedente_Click(sender As Object, e As EventArgs) Handles cmdLignePrecedente.Click
        Dim chaine As String

        If fichierUCI = "" Then
            If 0 < indexEPD Then
                indexEPD = indexEPD - 1
                lblInfo.Text = tabEPD(indexEPD)
                coup1 = ""
                pbEchiquier.Refresh()
                If indexEPD >= 0 And moteur_court <> "" Then
                    If tabEPD(indexEPD) <> "" Then
                        afficherEXP(tabEPD(indexEPD))
                    End If
                End If
            End If
        Else
            'partie precedente
            If 0 < indexPGN Then
                indexPGN = indexPGN - 1
                lstParties.SetSelected(indexPGN, True)
                If tabPGN(indexPGN) <> "" Then
                    tabCoups = Split(tabPGN(indexPGN), " ")
                    indexCoup = 0
                    If tabCoups(indexCoup) <> "" Then
                        chaine = tabCoups(indexCoup)
                        lblInfo.Text = "PGN " & (indexPGN + 1) & "/" & UBound(tabPGN) & " : " & chaine
                        ReDim tabEPD(0)
                        indexEPD = 0
                        tabEPD(indexEPD) = moteurEPD(lblEman.Text, chaine, tabStartpos(indexPGN))
                        coup1 = ""
                        If indexCoup < UBound(tabCoups) Then
                            coup1 = tabCoups(indexCoup + 1)
                        End If
                        pbEchiquier.Refresh()
                        If indexEPD >= 0 And moteur_court <> "" Then
                            If tabEPD(indexEPD) <> "" Then
                                afficherEXP(tabEPD(indexEPD))
                            End If
                        End If
                    End If
                End If
            End If
        End If

    End Sub

    Private Sub cmdCoupPrecedent_Click(sender As Object, e As EventArgs) Handles cmdCoupPrecedent.Click
        Dim chaine As String

        If 0 < indexCoup Then
            indexCoup = indexCoup - 1
            If fichierUCI = "" Then
                indexEPD = indexEPD - 1
                If tabCoups(indexCoup) <> "" Then
                    lstParties.SetSelected(indexEPD, True)
                    lblInfo.Text = lblInfo.Tag
                    If InStr(lblInfo.Tag, " ") > 0 Then
                        lblInfo.Tag.substring(0, lblInfo.Tag.indexof(" "))
                    End If
                    coup1 = ""
                    pbEchiquier.Refresh()
                    If indexEPD >= 0 And moteur_court <> "" Then
                        If tabEPD(indexEPD) <> "" Then
                            afficherEXP(tabEPD(indexEPD))
                        End If
                    End If
                End If
            Else
                If tabCoups(indexCoup) <> "" Then
                    chaine = lblInfo.Tag
                    lblInfo.Text = "PGN " & (indexPGN + 1) & "/" & UBound(tabPGN) & " : " & chaine
                    If InStr(lblInfo.Tag, " ") > 0 Then
                        lblInfo.Tag = chaine.Substring(0, chaine.LastIndexOf(" "))
                    End If
                    ReDim tabEPD(0)
                    indexEPD = 0
                    tabEPD(indexEPD) = moteurEPD(lblEman.Text, chaine, tabStartpos(indexPGN))
                    coup1 = ""
                    If indexCoup < UBound(tabCoups) Then
                        coup1 = tabCoups(indexCoup + 1)
                    End If
                    pbEchiquier.Refresh()
                    If indexEPD >= 0 And moteur_court <> "" Then
                        If tabEPD(indexEPD) <> "" Then
                            afficherEXP(tabEPD(indexEPD))
                        End If
                    End If
                End If
            End If

        End If
    End Sub

    Private Sub cmdCoupSuivant_Click(sender As Object, e As EventArgs) Handles cmdCoupSuivant.Click
        Dim chaine As String

        If indexCoup < UBound(tabCoups) Then
            indexCoup = indexCoup + 1
            If fichierUCI = "" Then
                indexEPD = indexEPD + 1
                If tabCoups(indexCoup) <> "" Then
                    lblInfo.Tag = lblInfo.Text
                    lblInfo.Text = Trim(lblInfo.Text & " " & tabCoups(indexCoup))
                    coup1 = ""
                    pbEchiquier.Refresh()
                    If indexEPD >= 0 And moteur_court <> "" Then
                        If tabEPD(indexEPD) <> "" Then
                            afficherEXP(tabEPD(indexEPD))
                        End If
                    End If
                End If
            Else
                If tabCoups(indexCoup) <> "" Then
                    lblInfo.Tag = lblInfo.Text.Substring(lblInfo.Text.IndexOf(":") + 2)
                    chaine = Trim(lblInfo.Tag & " " & tabCoups(indexCoup))
                    lblInfo.Text = "PGN " & (indexPGN + 1) & "/" & UBound(tabPGN) & " : " & chaine
                    ReDim tabEPD(0)
                    indexEPD = 0
                    tabEPD(indexEPD) = moteurEPD(lblEman.Text, chaine, tabStartpos(indexPGN))
                    coup1 = ""
                    If indexCoup < UBound(tabCoups) Then
                        coup1 = tabCoups(indexCoup + 1)
                    End If
                    pbEchiquier.Refresh()
                    If indexEPD >= 0 And moteur_court <> "" Then
                        If tabEPD(indexEPD) <> "" Then
                            afficherEXP(tabEPD(indexEPD))
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub cmdLigneSuivante_Click(sender As Object, e As EventArgs) Handles cmdLigneSuivante.Click
        Dim chaine As String

        If fichierUCI = "" Then
            If indexEPD < UBound(tabEPD) Then
                indexEPD = indexEPD + 1
                If tabEPD(indexEPD) <> "" Then
                    lstParties.SetSelected(indexEPD, True)
                    lblInfo.Text = tabEPD(indexEPD)
                    coup1 = ""
                    pbEchiquier.Refresh()
                    If indexEPD >= 0 And moteur_court <> "" Then
                        If tabEPD(indexEPD) <> "" Then
                            afficherEXP(tabEPD(indexEPD))
                        End If
                    End If
                End If
            End If
        Else
            'partie suivante
            If indexPGN < UBound(tabPGN) Then
                indexPGN = indexPGN + 1
                If tabPGN(indexPGN) <> "" Then
                    lstParties.SetSelected(indexPGN, True)
                    tabCoups = Split(tabPGN(indexPGN), " ")
                    indexCoup = 0
                    If tabCoups(indexCoup) <> "" Then
                        chaine = tabCoups(indexCoup)
                        lblInfo.Text = "PGN " & (indexPGN + 1) & "/" & UBound(tabPGN) & " : " & chaine
                        ReDim tabEPD(0)
                        indexEPD = 0
                        tabEPD(indexEPD) = moteurEPD(lblEman.Text, chaine, tabStartpos(indexPGN))
                        coup1 = ""
                        If indexCoup < UBound(tabCoups) Then
                            coup1 = tabCoups(indexCoup + 1)
                        End If
                        pbEchiquier.Refresh()
                        If indexEPD >= 0 And moteur_court <> "" Then
                            If tabEPD(indexEPD) <> "" Then
                                afficherEXP(tabEPD(indexEPD))
                            End If
                        End If
                    End If
                End If
            End If
        End If

    End Sub

    Private Sub txtRecherche_KeyUp(sender As Object, e As KeyEventArgs) Handles txtRecherche.KeyUp
        Dim tabTmp() As String
        If e.KeyCode = Keys.Enter Then
            Cursor = Cursors.WaitCursor
            If nbCaracteres(txtRecherche.Text, "/") = 7 _
            And (InStr(txtRecherche.Text, " w ") > 0 Or InStr(txtRecherche.Text, " b ") > 0) Then
                lblRecherche.Text = "Position EPD :"
                ReDim tabEPD(0)
                indexEPD = 0
                If InStr(txtRecherche.Text, " moves ", CompareMethod.Text) > 0 Then
                    tabTmp = Split(txtRecherche.Text, " moves ")
                    tabEPD(indexEPD) = moteurEPD(lblEman.Text, tabTmp(1), tabTmp(0))
                Else
                    tabEPD(indexEPD) = txtRecherche.Text
                End If
                lblInfo.Text = tabEPD(indexEPD)
                pbEchiquier.Refresh()
                If indexEPD >= 0 And moteur_court <> "" Then
                    If tabEPD(indexEPD) <> "" Then
                        afficherEXP(tabEPD(indexEPD))
                    End If
                End If
                If Me.Width <> 626 Then
                    Me.Width = 626
                    lstParties.Visible = False
                    Me.CenterToScreen()
                End If
            Else
                lblRecherche.Text = "Suite UCI :"
                ReDim tabEPD(0)
                indexEPD = 0
                tabEPD(indexEPD) = moteurEPD(lblEman.Text, txtRecherche.Text)
                lblInfo.Text = tabEPD(indexEPD)
                pbEchiquier.Refresh()
                If indexEPD >= 0 And moteur_court <> "" Then
                    If tabEPD(indexEPD) <> "" Then
                        afficherEXP(tabEPD(indexEPD))
                    End If
                End If
                If Me.Width <> 626 Then
                    Me.Width = 626
                    lstParties.Visible = False
                    Me.CenterToScreen()
                End If
            End If
            cmdLignePrecedente.Visible = False
            cmdCoupPrecedent.Visible = False
            cmdCoupSuivant.Visible = False
            cmdLigneSuivante.Visible = False
            Cursor = Cursors.Default
        End If
    End Sub

    Private Sub lstParties_DoubleClick(sender As Object, e As EventArgs) Handles lstParties.DoubleClick
        Dim chaine As String

        If fichierUCI = "" Then
            indexEPD = lstParties.SelectedIndex
            lblInfo.Text = tabEPD(indexEPD)
            pbEchiquier.Refresh()
            If indexEPD >= 0 And moteur_court <> "" Then
                If tabEPD(indexEPD) <> "" Then
                    afficherEXP(tabEPD(indexEPD))
                End If
            End If
        Else
            indexPGN = lstParties.SelectedIndex
            If tabPGN(indexPGN) <> "" Then
                tabCoups = Split(tabPGN(indexPGN), " ")
                indexCoup = 0
                If tabCoups(indexCoup) <> "" Then
                    chaine = tabCoups(indexCoup)
                    lblInfo.Text = "PGN " & (indexPGN + 1) & "/" & UBound(tabPGN) & " : " & chaine
                    ReDim tabEPD(0)
                    indexEPD = 0
                    tabEPD(indexEPD) = moteurEPD(lblEman.Text, chaine, tabStartpos(indexPGN))
                    pbEchiquier.Refresh()
                    If indexEPD >= 0 And moteur_court <> "" Then
                        If tabEPD(indexEPD) <> "" Then
                            afficherEXP(tabEPD(indexEPD))
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub cmdAnalyse_Click(sender As Object, e As EventArgs) Handles cmdAnalyse.Click
        analyse = False
        frmOptions.ShowDialog()
        If Not analyse Then
            Exit Sub
        End If

        My.Computer.FileSystem.WriteAllText(fichierINI, "moteurEXP = " & lblEman.Text & vbCrLf _
                                                      & "fichierEXP = " & lblExperience.Text & vbCrLf _
                                                      & "profMinAnalyse = " & profMinAnalyse & vbCrLf _
                                                      & "profMaxAnalyse = " & profMaxAnalyse & vbCrLf _
                                                      & "modeAnalyse = " & modeAnalyse, False)

        cmdAnalyse.Visible = False
        cmdDefrag.Visible = False
        cmdStop.Visible = True
        txtRecherche.Enabled = False
        cmdFichier.Enabled = False
        cmdMoteur.Enabled = False
        cmdRecherche.Enabled = False
        lstParties.Enabled = False

        timerAnalyse.Enabled = True
        tacheAnalyse = New Thread(AddressOf analysePosition)
        tacheAnalyse.Start()

    End Sub

    Private Sub timerAnalyse_Tick(sender As Object, e As EventArgs) Handles timerAnalyse.Tick
        If Not cmdStop.Visible Then
            timerAnalyse.Enabled = False
            coup1 = ""
            pbEchiquier.Refresh()

            cmdAnalyse.Visible = True
            cmdDefrag.Visible = True
            txtRecherche.Enabled = True
            cmdFichier.Enabled = True
            cmdMoteur.Enabled = True
            cmdRecherche.Enabled = True
            lstParties.Enabled = True

            Me.Text = My.Application.Info.AssemblyName
            Application.DoEvents()
        End If
        
    End Sub

    Private Function dejaAnalysee(coup As String, profDemandee As Integer) As Boolean
        Dim i As Integer, trouve As Boolean, tabChaine() As String, prof As Integer
        Dim tabTmp() As String

        trouve = False
        tabChaine = Split(listeExperience, vbCrLf)
        For i = 0 To UBound(tabChaine)
            If tabChaine(i) <> "" Then
                tabTmp = Split(Replace(tabChaine(i), ":", ","), ",")
                If coup = Trim(tabTmp(1)) Then
                    prof = CInt(Trim(tabTmp(3)))
                    If prof >= profDemandee Then
                        trouve = True
                        Exit For
                    End If
                End If
            End If
        Next

        Return trouve
    End Function

    Private Sub analysePosition()
        Dim chaine As String, chaineInfo As String, liste As String, tabChaine() As String
        Dim prof As Integer, position As String, tabTmp() As String
        Dim entryStored As Boolean, tabEXP(0) As Byte, scoreCP As Integer

        If cmdStop.Visible Then
            position = tabEPD(indexEPD)

            entree.WriteLine("position fen " & position)

            liste = ""
            Select Case modeAnalyse
                Case 1
                    'on cherche tous les coups possibles
                    entree.WriteLine("setoption name MultiPV value " & maxMultiPVMoteur(moteur_court))
                    entree.WriteLine("go depth 1")

                    chaine = ""
                    While InStr(chaine, "bestmove", CompareMethod.Text) = 0
                        chaine = sortie.ReadLine
                        If InStr(chaine, " pv ", CompareMethod.Text) > 0 Then
                            tabChaine = Split(chaine, " ")
                            For i = 0 To UBound(tabChaine) - 1
                                If InStr(tabChaine(i), "pv", CompareMethod.Text) > 0 _
                                And tabChaine(i + 1) <> "" And Len(tabChaine(i + 1)) = 4 Then
                                    liste = liste & tabChaine(i + 1) & vbCrLf
                                End If
                            Next
                        End If
                        Threading.Thread.Sleep(1)
                    End While
                    entree.WriteLine("stop")

                    entree.WriteLine("isready")
                    chaine = ""
                    While InStr(chaine, "readyok") = 0
                        chaine = sortie.ReadLine
                        Threading.Thread.Sleep(1)
                    End While

                Case 2 'only played moves
                    tabChaine = Split(listeExperience, vbCrLf)
                    For i = 1 To 20
                        If i < tabChaine.Length Then
                            If tabChaine(i - 1) <> "" Then
                                tabTmp = Split(Replace(tabChaine(i - 1), ":", ","), ",")
                                liste = liste & Trim(tabTmp(1)) & vbCrLf
                            End If
                        End If
                    Next

                Case 3 'only moves with positive score
                    tabChaine = Split(listeExperience, vbCrLf)
                    For i = 1 To 20
                        If i < tabChaine.Length Then
                            If tabChaine(i - 1) <> "" Then
                                tabTmp = Split(Replace(tabChaine(i - 1), ":", ","), ",")
                                If InStr(tabTmp(5), "-") = 0 Then
                                    liste = liste & Trim(tabTmp(1)) & vbCrLf
                                End If
                            End If
                        End If
                    Next

                Case 4 'search bestmove
                    liste = liste & vbCrLf

                Case Else

            End Select

            'on analyse les coups les uns après les autres
            entree.WriteLine("setoption name MultiPV value 1")
            tabChaine = Split(trierChaine(liste, vbCrLf), vbCrLf)
            prof = profMinAnalyse
            chaineInfo = ""
            While cmdStop.Visible And prof <= profMaxAnalyse
                For i = 0 To UBound(tabChaine)
                    If (modeAnalyse = 4 Or tabChaine(i) <> "") And Not dejaAnalysee(tabChaine(i), prof) Then
                        If modeAnalyse = 4 Then
                            entree.WriteLine("go depth " & prof)
                            If chaineInfo <> "" Then
                                coup1 = Trim(chaineInfo.Substring(0, chaineInfo.IndexOf(" ")))
                            End If
                        Else
                            entree.WriteLine("go depth " & prof & " searchmoves " & tabChaine(i))
                            coup1 = tabChaine(i)
                        End If

                        pbEchiquier.Refresh()

                        chaine = ""
                        chaineInfo = ""
                        While cmdStop.Visible And InStr(chaine, "bestmove", CompareMethod.Text) = 0
                            chaine = sortie.ReadLine
                            If InStr(chaine, " pv ", CompareMethod.Text) > 0 And InStr(chaine, "upperbound", CompareMethod.Text) = 0 And InStr(chaine, "lowerbound", CompareMethod.Text) = 0 And InStr(chaine, " depth " & prof, CompareMethod.Text) > 0 Then
                                chaineInfo = formaterInfos(chaine, position)
                                If InStr(chaineInfo, "mate", CompareMethod.Text) > 0 Then
                                    chaineInfo = Replace(chaineInfo, "{mate -", "{-M")
                                    chaineInfo = Replace(chaineInfo, "{mate ", "{+M")
                                End If
                                Me.Text = "Prof " & Format(prof) & "/" & profMaxAnalyse & ", coup " & Format(i + 1) & "/" & tabChaine.Length & " : " & chaineInfo
                            End If
                            Threading.Thread.Sleep(1)
                        End While
                        entree.WriteLine("stop")

                        entree.WriteLine("isready")
                        chaine = ""
                        While InStr(chaine, "readyok") = 0
                            chaine = sortie.ReadLine
                            Threading.Thread.Sleep(1)
                        End While
                    End If
                Next
                entryStored = False
                entree.WriteLine("ucinewgame")

                entree.WriteLine("isready")
                chaine = ""
                While InStr(chaine, "readyok") = 0
                    chaine = sortie.ReadLine
                    If InStr(chaine, "Saved", CompareMethod.Text) > 0 And InStr(chaine, "entries", CompareMethod.Text) > 0 Then
                        entryStored = True
                    End If
                    Threading.Thread.Sleep(1)
                End While

                If Not entryStored Then
                    'obtenir la position fen puis voir plainToEXP pour la structure d'une entrée dans fichierEXP
                    If InStr(moteur_court, "eman 6", CompareMethod.Text) > 0 Then
                        ReDim tabEXP(31)
                    ElseIf InStr(moteur_court, "eman 7", CompareMethod.Text) > 0 Or InStr(moteur_court, "eman 8", CompareMethod.Text) > 0 _
                        Or InStr(moteur_court, "hypnos", CompareMethod.Text) > 0 Or InStr(moteur_court, "stockfishmz", CompareMethod.Text) > 0 _
                        Or InStr(moteur_court, "aurora", CompareMethod.Text) > 0 Then
                        ReDim tabEXP(23)
                    End If

                    'a1b1 {1.59/22 2} => a1b1
                    coup1 = Trim(chaineInfo.Substring(0, chaineInfo.IndexOf(" ")))

                    'a1b1 {1.59/22 2} => 1.59/22 2}
                    chaine = Replace(chaineInfo, coup1 & " {", "")
                    '1.59/22 2} => 1.59
                    chaine = chaine.Substring(0, chaine.IndexOf("/"))
                    '1.59 => 159
                    scoreCP = CSng(Replace(chaine, ".", ",")) * 100

                    'position 2br1rk1/pp3ppp/1qpn1b2/2N5/3P4/P2BBQ2/1P3PPP/R3R1K1 w - - 0 20 => 540B1223AA5EEE2B => 2BEE5EAA23120B54
                    'coup1 a1b1 => 0 000 000(1) 000(a) 000(1) 001(b) => 00 01 => 01 00
                    'scoreCP 159 => 01 3E => 3E 01
                    'D22 => 16

                    'position                coup  ?  ?  score       profondeur  visites
                    '0  1  2  3  4  5  6  7  8  9  a  b  c  d  e  f  0  1  2  3  4  5  6  7
                    '2B EE 5E AA 23 12 0B 54 01 00 00 00 C8 01 00 00 16 00 00 00 01 00 00 00

                    entreeEXP(tabEXP, position, coup1, scoreCP, 100, prof, 1, entree, sortie, moteur_court)

                    My.Computer.FileSystem.WriteAllBytes(lblExperience.Text, tabEXP, True)
                End If

                prof = prof + 2

                afficherEXP(position)
            End While
            coup1 = ""

            cmdStop_Click(Nothing, Nothing)

            pbEchiquier.Refresh()
            If indexEPD >= 0 And moteur_court <> "" Then
                If position <> "" Then
                    afficherEXP(position)
                End If
            End If
        End If
    End Sub

    Private Sub cmdStop_Click(sender As Object, e As EventArgs) Handles cmdStop.Click
        Dim chaine As String

        Me.Tag = Me.Text

        chaine = ""
        While InStr(chaine, "readyok") = 0
            entree.WriteLine("stop")
            entree.WriteLine("isready")
            chaine = sortie.ReadLine
            Threading.Thread.Sleep(1)
        End While

        Me.Text = "Loading " & nomFichier(lblExperience.Text) & "..."
        tabExperience = My.Computer.FileSystem.ReadAllBytes(lblExperience.Text)

        cmdStop.Visible = False
        Me.Text = Me.Tag
    End Sub

    Private Sub cmdDefrag_Click(sender As Object, e As EventArgs) Handles cmdDefrag.Click
        Me.Tag = Me.Text

        cmdAnalyse.Enabled = False
        cmdFichier.Enabled = False
        cmdDefrag.Enabled = False
        txtRecherche.Enabled = False
        cmdMoteur.Enabled = False
        cmdRecherche.Enabled = False

        Me.Text = "Defrag " & nomFichier(lblExperience.Text) & "..."

        If InStr(moteur_court, "eman 6", CompareMethod.Text) > 0 And InStr(entete, "Collisions: 0", CompareMethod.Text) = 0 Then
            entete = defragEXP(lblExperience.Text, 1, entree, sortie)
        ElseIf (InStr(moteur_court, "eman 7", CompareMethod.Text) > 0 Or InStr(moteur_court, "eman 8", CompareMethod.Text) > 0) And InStr(entete, "Duplicate moves: 0", CompareMethod.Text) = 0 Then
            entete = defragEXP(lblExperience.Text, 4, entree, sortie)
        ElseIf (InStr(moteur_court, "hypnos", CompareMethod.Text) > 0 Or InStr(moteur_court, "stockfishmz", CompareMethod.Text) > 0 Or InStr(moteur_court, "aurora", CompareMethod.Text) > 0) And InStr(entete, "Duplicate moves: 0", CompareMethod.Text) = 0 Then
            entete = defragHypnos(lblEman.Text, lblExperience.Text, entree, sortie)
        End If
        lblInfo.Text = entete

        Me.Text = "Loading " & nomFichier(lblExperience.Text) & "..."
        tabExperience = My.Computer.FileSystem.ReadAllBytes(lblExperience.Text)

        pbEchiquier.Refresh()
        If indexEPD >= 0 And moteur_court <> "" Then
            If tabEPD(indexEPD) <> "" Then
                afficherEXP(tabEPD(indexEPD))
            End If
        End If

        cmdAnalyse.Enabled = True
        cmdFichier.Enabled = True
        cmdDefrag.Enabled = True
        txtRecherche.Enabled = True
        cmdMoteur.Enabled = True
        cmdRecherche.Enabled = True

        Me.Text = Me.Tag
        Application.DoEvents()
    End Sub

    Private Sub lblRang1_DoubleClick(sender As Object, e As EventArgs) Handles lblRang1.DoubleClick
        supCoup(entree, sortie, tabEPD(indexEPD), lblCoup1.Text)
    End Sub

    Private Sub lblRang2_DoubleClick(sender As Object, e As EventArgs) Handles lblRang2.DoubleClick
        supCoup(entree, sortie, tabEPD(indexEPD), lblCoup2.Text)
    End Sub

    Private Sub lblRang3_DoubleClick(sender As Object, e As EventArgs) Handles lblRang3.DoubleClick
        supCoup(entree, sortie, tabEPD(indexEPD), lblCoup3.Text)
    End Sub

    Private Sub lblRang4_DoubleClick(sender As Object, e As EventArgs) Handles lblRang4.DoubleClick
        supCoup(entree, sortie, tabEPD(indexEPD), lblCoup4.Text)
    End Sub

    Private Sub lblRang5_DoubleClick(sender As Object, e As EventArgs) Handles lblRang5.DoubleClick
        supCoup(entree, sortie, tabEPD(indexEPD), lblCoup5.Text)
    End Sub

    Private Sub lblRang6_DoubleClick(sender As Object, e As EventArgs) Handles lblRang6.DoubleClick
        supCoup(entree, sortie, tabEPD(indexEPD), lblCoup6.Text)
    End Sub

    Private Sub lblRang7_DoubleClick(sender As Object, e As EventArgs) Handles lblRang7.DoubleClick
        supCoup(entree, sortie, tabEPD(indexEPD), lblCoup7.Text)
    End Sub

    Private Sub lblRang8_DoubleClick(sender As Object, e As EventArgs) Handles lblRang8.DoubleClick
        supCoup(entree, sortie, tabEPD(indexEPD), lblCoup8.Text)
    End Sub

    Private Sub lblRang9_DoubleClick(sender As Object, e As EventArgs) Handles lblRang9.DoubleClick
        supCoup(entree, sortie, tabEPD(indexEPD), lblCoup9.Text)
    End Sub

    Private Sub lblRang10_DoubleClick(sender As Object, e As EventArgs) Handles lblRang10.DoubleClick
        supCoup(entree, sortie, tabEPD(indexEPD), lblCoup10.Text)
    End Sub

    Private Sub lblRang11_DoubleClick(sender As Object, e As EventArgs) Handles lblRang11.DoubleClick
        supCoup(entree, sortie, tabEPD(indexEPD), lblCoup11.Text)
    End Sub

    Private Sub lblRang12_DoubleClick(sender As Object, e As EventArgs) Handles lblRang12.DoubleClick
        supCoup(entree, sortie, tabEPD(indexEPD), lblCoup12.Text)
    End Sub

    Private Sub lblRang13_DoubleClick(sender As Object, e As EventArgs) Handles lblRang13.DoubleClick
        supCoup(entree, sortie, tabEPD(indexEPD), lblCoup13.Text)
    End Sub

    Private Sub lblRang14_DoubleClick(sender As Object, e As EventArgs) Handles lblRang14.DoubleClick
        supCoup(entree, sortie, tabEPD(indexEPD), lblCoup14.Text)
    End Sub

    Private Sub lblRang15_DoubleClick(sender As Object, e As EventArgs) Handles lblRang15.DoubleClick
        supCoup(entree, sortie, tabEPD(indexEPD), lblCoup15.Text)
    End Sub

    Private Sub lblRang16_DoubleClick(sender As Object, e As EventArgs) Handles lblRang16.DoubleClick
        supCoup(entree, sortie, tabEPD(indexEPD), lblCoup16.Text)
    End Sub

    Private Sub lblRang17_DoubleClick(sender As Object, e As EventArgs) Handles lblRang17.DoubleClick
        supCoup(entree, sortie, tabEPD(indexEPD), lblCoup17.Text)
    End Sub

    Private Sub lblRang18_DoubleClick(sender As Object, e As EventArgs) Handles lblRang18.DoubleClick
        supCoup(entree, sortie, tabEPD(indexEPD), lblCoup18.Text)
    End Sub

    Private Sub lblRang19_DoubleClick(sender As Object, e As EventArgs) Handles lblRang19.DoubleClick
        supCoup(entree, sortie, tabEPD(indexEPD), lblCoup19.Text)
    End Sub

    Private Sub lblRang20_DoubleClick(sender As Object, e As EventArgs) Handles lblRang20.DoubleClick
        supCoup(entree, sortie, tabEPD(indexEPD), lblCoup20.Text)
    End Sub

    Private Sub supCoup(entreeRang As System.IO.StreamWriter, sortieRang As System.IO.StreamReader, epd As String, coup As String)
        Dim critereEXP(23) As Byte, i As Long, critereEXPbis(23) As Byte

        If coup = "" Then
            Exit Sub
        End If
        Cursor = Cursors.WaitCursor
        txtRecherche.Enabled = False

        'position                coup   ?   ?    score             profondeur        visites
        '0 1  2  3  4  5  6  7   8  9   a   b    c   d   e   f     0   1   2   3     4   5   6   7
        '0 +1 +2 +3 +4 +5 +6 +7  +8 +9  +10 +11  +12 +13 +14 +15   +16 +17 +18 +19   +20 +21 +22 +23
        'tabExperience(i+16) = 60 (profondeur)
        'tabExperience(i+20) = 120 (visites)

        If MsgBox("Delete '" & coup & "' ?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            'on prépare un critère de recherche basé sur la position en cours et le coup sélectionné
            Array.Copy(epdToEXP(entreeRang, sortieRang, epd), 0, critereEXP, 0, 8)
            Array.Copy(moveToEXP(coup), 0, critereEXP, 8, 2)
            Array.Copy(moveToEXP(coup, True), 0, critereEXPbis, 8, 2)

            'on décharge le moteur
            dechargerMoteur()
            Application.DoEvents()

            'on cherche la position et le coup dans Eman.exp
            '0-25 = SugaR Experience version 2
            For i = 26 To UBound(tabExperience) Step 24
                'position
                If tabExperience(i) = critereEXP(0) Then
                    If tabExperience(i + 1) = critereEXP(1) Then
                        If tabExperience(i + 2) = critereEXP(2) Then
                            If tabExperience(i + 3) = critereEXP(3) Then
                                If tabExperience(i + 4) = critereEXP(4) Then
                                    If tabExperience(i + 5) = critereEXP(5) Then
                                        If tabExperience(i + 6) = critereEXP(6) Then
                                            If tabExperience(i + 7) = critereEXP(7) Then
                                                'coup
                                                If (tabExperience(i + 8) = critereEXP(8) And tabExperience(i + 9) = critereEXP(9)) _
                                                Or (tabExperience(i + 8) = critereEXPbis(8) And tabExperience(i + 9) = critereEXPbis(9)) Then
                                                    'pour effacer ce coup,
                                                    If i >= 24 Then
                                                        'on l'écrase avec les données de celui d'avant
                                                        For j = 0 To 23
                                                            tabExperience(i + j) = tabExperience(i - 24 + j)
                                                        Next
                                                    Else
                                                        'ou celui d'après
                                                        For j = 0 To 23
                                                            tabExperience(i + j) = tabExperience(i + 24 + j)
                                                        Next
                                                    End If
                                                    My.Computer.FileSystem.WriteAllBytes(lblExperience.Text, tabExperience, False)
                                                    'dans tous les cas on sort ! (Sinon si le fichier n'était pas défragmenté, ça peut boucler très très longtemps)
                                                    Exit For
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next

            'on recharge le moteur
            chargerMoteur(lblEman.Text, lblExperience.Text)

            'on met à jour l'affichage
            If indexEPD >= 0 And moteur_court <> "" Then
                If tabEPD(indexEPD) <> "" Then
                    afficherEXP(tabEPD(indexEPD))
                End If
            End If

        End If

        txtRecherche.Enabled = True
        Cursor = Cursors.Default
    End Sub

    Private Sub lblCoup1_Click(sender As Object, e As EventArgs) Handles lblCoup1.Click
        majRecherche(lblCoup1.Text)
    End Sub

    Private Sub lblCoup2_Click(sender As Object, e As EventArgs) Handles lblCoup2.Click
        majRecherche(lblCoup2.Text)
    End Sub

    Private Sub lblCoup3_Click(sender As Object, e As EventArgs) Handles lblCoup3.Click
        majRecherche(lblCoup3.Text)
    End Sub

    Private Sub lblCoup4_Click(sender As Object, e As EventArgs) Handles lblCoup4.Click
        majRecherche(lblCoup4.Text)
    End Sub

    Private Sub lblCoup5_Click(sender As Object, e As EventArgs) Handles lblCoup5.Click
        majRecherche(lblCoup5.Text)
    End Sub

    Private Sub lblCoup6_Click(sender As Object, e As EventArgs) Handles lblCoup6.Click
        majRecherche(lblCoup6.Text)
    End Sub

    Private Sub lblCoup7_Click(sender As Object, e As EventArgs) Handles lblCoup7.Click
        majRecherche(lblCoup7.Text)
    End Sub

    Private Sub lblCoup8_Click(sender As Object, e As EventArgs) Handles lblCoup8.Click
        majRecherche(lblCoup8.Text)
    End Sub

    Private Sub lblCoup9_Click(sender As Object, e As EventArgs) Handles lblCoup9.Click
        majRecherche(lblCoup9.Text)
    End Sub

    Private Sub lblCoup10_Click(sender As Object, e As EventArgs) Handles lblCoup10.Click
        majRecherche(lblCoup10.Text)
    End Sub

    Private Sub lblCoup11_Click(sender As Object, e As EventArgs) Handles lblCoup11.Click
        majRecherche(lblCoup11.Text)
    End Sub

    Private Sub lblCoup12_Click(sender As Object, e As EventArgs) Handles lblCoup12.Click
        majRecherche(lblCoup12.Text)
    End Sub

    Private Sub lblCoup13_Click(sender As Object, e As EventArgs) Handles lblCoup13.Click
        majRecherche(lblCoup13.Text)
    End Sub

    Private Sub lblCoup14_Click(sender As Object, e As EventArgs) Handles lblCoup14.Click
        majRecherche(lblCoup14.Text)
    End Sub

    Private Sub lblCoup15_Click(sender As Object, e As EventArgs) Handles lblCoup15.Click
        majRecherche(lblCoup15.Text)
    End Sub

    Private Sub lblCoup16_Click(sender As Object, e As EventArgs) Handles lblCoup16.Click
        majRecherche(lblCoup16.Text)
    End Sub

    Private Sub lblCoup17_Click(sender As Object, e As EventArgs) Handles lblCoup17.Click
        majRecherche(lblCoup17.Text)
    End Sub

    Private Sub lblCoup18_Click(sender As Object, e As EventArgs) Handles lblCoup18.Click
        majRecherche(lblCoup18.Text)
    End Sub

    Private Sub lblCoup19_Click(sender As Object, e As EventArgs) Handles lblCoup19.Click
        majRecherche(lblCoup19.Text)
    End Sub

    Private Sub lblCoup20_Click(sender As Object, e As EventArgs) Handles lblCoup20.Click
        majRecherche(lblCoup20.Text)
    End Sub

    Private Sub majRecherche(coupUCI As String)
        If Not cmdLignePrecedente.Visible And Not cmdLigneSuivante.Visible And coupUCI <> "" Then
            Cursor = Cursors.WaitCursor
            txtRecherche.Enabled = False
            ReDim tabEPD(0)
            indexEPD = 0
            If txtRecherche.Text <> "" And InStr(lblRecherche.Text, "EPD", CompareMethod.Text) > 0 Then
                If InStr(txtRecherche.Text, "moves", CompareMethod.Text) > 0 Then
                    txtRecherche.Text = Trim(txtRecherche.Text & " " & coupUCI)
                Else
                    txtRecherche.Text = Trim(txtRecherche.Text & " moves " & coupUCI)
                End If
                tabEPD(indexEPD) = moteurEPD(lblEman.Text, txtRecherche.Text.Substring(txtRecherche.Text.IndexOf(" moves ") + 7), txtRecherche.Text.Substring(0, txtRecherche.Text.IndexOf(" moves ")))
            Else
                txtRecherche.Text = Trim(txtRecherche.Text & " " & coupUCI)
                lblRecherche.Text = "Suite UCI :"
                tabEPD(indexEPD) = moteurEPD(lblEman.Text, txtRecherche.Text)
            End If

            lblInfo.Text = tabEPD(indexEPD)
            pbEchiquier.Refresh()
            If indexEPD >= 0 And moteur_court <> "" Then
                If tabEPD(indexEPD) <> "" Then
                    afficherEXP(tabEPD(indexEPD))
                End If
            End If
            txtRecherche.Enabled = True
            Cursor = Cursors.Default
        End If
    End Sub

    Private Sub lblScore1_DoubleClick(sender As Object, e As EventArgs) Handles lblScore1.DoubleClick
        majProfondeur(entree, sortie, tabEPD(indexEPD), lblCoup1.Text, lblScore1.Text)
    End Sub

    Private Sub lblScore2_DoubleClick(sender As Object, e As EventArgs) Handles lblScore2.DoubleClick
        majProfondeur(entree, sortie, tabEPD(indexEPD), lblCoup2.Text, lblScore2.Text)
    End Sub

    Private Sub lblScore3_DoubleClick(sender As Object, e As EventArgs) Handles lblScore3.DoubleClick
        majProfondeur(entree, sortie, tabEPD(indexEPD), lblCoup3.Text, lblScore3.Text)
    End Sub

    Private Sub lblScore4_DoubleClick(sender As Object, e As EventArgs) Handles lblScore4.DoubleClick
        majProfondeur(entree, sortie, tabEPD(indexEPD), lblCoup4.Text, lblScore4.Text)
    End Sub

    Private Sub lblScore5_DoubleClick(sender As Object, e As EventArgs) Handles lblScore5.DoubleClick
        majProfondeur(entree, sortie, tabEPD(indexEPD), lblCoup5.Text, lblScore5.Text)
    End Sub

    Private Sub lblScore6_DoubleClick(sender As Object, e As EventArgs) Handles lblScore6.DoubleClick
        majProfondeur(entree, sortie, tabEPD(indexEPD), lblCoup6.Text, lblScore6.Text)
    End Sub

    Private Sub lblScore7_DoubleClick(sender As Object, e As EventArgs) Handles lblScore7.DoubleClick
        majProfondeur(entree, sortie, tabEPD(indexEPD), lblCoup7.Text, lblScore7.Text)
    End Sub

    Private Sub lblScore8_DoubleClick(sender As Object, e As EventArgs) Handles lblScore8.DoubleClick
        majProfondeur(entree, sortie, tabEPD(indexEPD), lblCoup8.Text, lblScore8.Text)
    End Sub

    Private Sub lblScore9_DoubleClick(sender As Object, e As EventArgs) Handles lblScore9.DoubleClick
        majProfondeur(entree, sortie, tabEPD(indexEPD), lblCoup9.Text, lblScore9.Text)
    End Sub

    Private Sub lblScore10_DoubleClick(sender As Object, e As EventArgs) Handles lblScore10.DoubleClick
        majProfondeur(entree, sortie, tabEPD(indexEPD), lblCoup10.Text, lblScore10.Text)
    End Sub

    Private Sub lblScore11_DoubleClick(sender As Object, e As EventArgs) Handles lblScore11.DoubleClick
        majProfondeur(entree, sortie, tabEPD(indexEPD), lblCoup11.Text, lblScore11.Text)
    End Sub

    Private Sub lblScore12_DoubleClick(sender As Object, e As EventArgs) Handles lblScore12.DoubleClick
        majProfondeur(entree, sortie, tabEPD(indexEPD), lblCoup12.Text, lblScore12.Text)
    End Sub

    Private Sub lblScore13_DoubleClick(sender As Object, e As EventArgs) Handles lblScore13.DoubleClick
        majProfondeur(entree, sortie, tabEPD(indexEPD), lblCoup13.Text, lblScore13.Text)
    End Sub

    Private Sub lblScore14_DoubleClick(sender As Object, e As EventArgs) Handles lblScore14.DoubleClick
        majProfondeur(entree, sortie, tabEPD(indexEPD), lblCoup14.Text, lblScore14.Text)
    End Sub

    Private Sub lblScore15_DoubleClick(sender As Object, e As EventArgs) Handles lblScore15.DoubleClick
        majProfondeur(entree, sortie, tabEPD(indexEPD), lblCoup15.Text, lblScore15.Text)
    End Sub

    Private Sub lblScore16_DoubleClick(sender As Object, e As EventArgs) Handles lblScore16.DoubleClick
        majProfondeur(entree, sortie, tabEPD(indexEPD), lblCoup16.Text, lblScore16.Text)
    End Sub

    Private Sub lblScore17_DoubleClick(sender As Object, e As EventArgs) Handles lblScore17.DoubleClick
        majProfondeur(entree, sortie, tabEPD(indexEPD), lblCoup17.Text, lblScore17.Text)
    End Sub

    Private Sub lblScore18_DoubleClick(sender As Object, e As EventArgs) Handles lblScore18.DoubleClick
        majProfondeur(entree, sortie, tabEPD(indexEPD), lblCoup18.Text, lblScore18.Text)
    End Sub

    Private Sub lblScore19_DoubleClick(sender As Object, e As EventArgs) Handles lblScore19.DoubleClick
        majProfondeur(entree, sortie, tabEPD(indexEPD), lblCoup19.Text, lblScore19.Text)
    End Sub

    Private Sub lblScore20_DoubleClick(sender As Object, e As EventArgs) Handles lblScore20.DoubleClick
        majProfondeur(entree, sortie, tabEPD(indexEPD), lblCoup20.Text, lblScore20.Text)
    End Sub

    Private Sub majProfondeur(entreeVisite As System.IO.StreamWriter, sortieVisite As System.IO.StreamReader, epd As String, coup As String, scoreDepth As String)
        Dim critereEXP(23) As Byte, i As Long, nouvelleProf As Integer, reponse As String, profActuelle As Integer

        If coup = "" Then
            Exit Sub
        End If
        Cursor = Cursors.WaitCursor
        txtRecherche.Enabled = False

        profActuelle = CInt(Replace(scoreDepth.Substring(scoreDepth.IndexOf("/") + 1), "}", ""))

        'position                coup   ?   ?    score             profondeur        visites
        '0 1  2  3  4  5  6  7   8  9   a   b    c   d   e   f     0   1   2   3     4   5   6   7
        '0 +1 +2 +3 +4 +5 +6 +7  +8 +9  +10 +11  +12 +13 +14 +15   +16 +17 +18 +19   +20 +21 +22 +23
        'tabExperience(i+16) = 60 (profondeur)
        'tabExperience(i+20) = 120 (visites)

        reponse = InputBox("New depth value ?", "min 4 <= depth value <= max 245", profActuelle)
        If reponse <> "" Then
            If IsNumeric(reponse) Then
                nouvelleProf = CInt(reponse)
                If 4 <= nouvelleProf And nouvelleProf <= 245 And nouvelleProf <> profActuelle Then
                    'on prépare un critère de recherche basé sur la position en cours et le coup sélectionné
                    Array.Copy(epdToEXP(entreeVisite, sortieVisite, epd), 0, critereEXP, 0, 8)
                    Array.Copy(moveToEXP(coup), 0, critereEXP, 8, 2)

                    'on décharge le moteur
                    dechargerMoteur()
                    Application.DoEvents()

                    'on met à jour la profondeur dans Eman.exp
                    '0-25 = SugaR Experience version 2
                    For i = 26 To UBound(tabExperience) Step 24
                        'position
                        If tabExperience(i) = critereEXP(0) Then
                            If tabExperience(i + 1) = critereEXP(1) Then
                                If tabExperience(i + 2) = critereEXP(2) Then
                                    If tabExperience(i + 3) = critereEXP(3) Then
                                        If tabExperience(i + 4) = critereEXP(4) Then
                                            If tabExperience(i + 5) = critereEXP(5) Then
                                                If tabExperience(i + 6) = critereEXP(6) Then
                                                    If tabExperience(i + 7) = critereEXP(7) Then
                                                        'coup
                                                        If tabExperience(i + 8) = critereEXP(8) And tabExperience(i + 9) = critereEXP(9) Then
                                                            'profondeur
                                                            If tabExperience(i + 16) <> nouvelleProf Then
                                                                tabExperience(i + 16) = nouvelleProf
                                                                My.Computer.FileSystem.WriteAllBytes(lblExperience.Text, tabExperience, False)
                                                                'dans tous les cas on sort ! (Sinon si le fichier n'était pas défragmenté, ça peut boucler très très longtemps)
                                                                Exit For
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next

                    'on recharge le moteur
                    chargerMoteur(lblEman.Text, lblExperience.Text)

                    'on met à jour l'affichage
                    If indexEPD >= 0 And moteur_court <> "" Then
                        If tabEPD(indexEPD) <> "" Then
                            afficherEXP(tabEPD(indexEPD))
                        End If
                    End If

                End If
            End If
        End If
        
        txtRecherche.Enabled = True
        Cursor = Cursors.Default
    End Sub

    Private Sub lblVisite1_DoubleClick(sender As Object, e As EventArgs) Handles lblVisite1.DoubleClick
        majVisite(entree, sortie, tabEPD(indexEPD), lblCoup1.Text, lblVisite1.Text)
    End Sub

    Private Sub lblVisite2_DoubleClick(sender As Object, e As EventArgs) Handles lblVisite2.DoubleClick
        majVisite(entree, sortie, tabEPD(indexEPD), lblCoup2.Text, lblVisite2.Text)
    End Sub

    Private Sub lblVisite3_DoubleClick(sender As Object, e As EventArgs) Handles lblVisite3.DoubleClick
        majVisite(entree, sortie, tabEPD(indexEPD), lblCoup3.Text, lblVisite3.Text)
    End Sub

    Private Sub lblVisite4_DoubleClick(sender As Object, e As EventArgs) Handles lblVisite4.DoubleClick
        majVisite(entree, sortie, tabEPD(indexEPD), lblCoup4.Text, lblVisite4.Text)
    End Sub

    Private Sub lblVisite5_DoubleClick(sender As Object, e As EventArgs) Handles lblVisite5.DoubleClick
        majVisite(entree, sortie, tabEPD(indexEPD), lblCoup5.Text, lblVisite5.Text)
    End Sub

    Private Sub lblVisite6_DoubleClick(sender As Object, e As EventArgs) Handles lblVisite6.DoubleClick
        majVisite(entree, sortie, tabEPD(indexEPD), lblCoup6.Text, lblVisite6.Text)
    End Sub

    Private Sub lblVisite7_DoubleClick(sender As Object, e As EventArgs) Handles lblVisite7.DoubleClick
        majVisite(entree, sortie, tabEPD(indexEPD), lblCoup7.Text, lblVisite7.Text)
    End Sub

    Private Sub lblVisite8_DoubleClick(sender As Object, e As EventArgs) Handles lblVisite8.DoubleClick
        majVisite(entree, sortie, tabEPD(indexEPD), lblCoup8.Text, lblVisite8.Text)
    End Sub

    Private Sub lblVisite9_DoubleClick(sender As Object, e As EventArgs) Handles lblVisite9.DoubleClick
        majVisite(entree, sortie, tabEPD(indexEPD), lblCoup9.Text, lblVisite9.Text)
    End Sub

    Private Sub lblVisite10_DoubleClick(sender As Object, e As EventArgs) Handles lblVisite10.DoubleClick
        majVisite(entree, sortie, tabEPD(indexEPD), lblCoup10.Text, lblVisite10.Text)
    End Sub

    Private Sub lblVisite11_DoubleClick(sender As Object, e As EventArgs) Handles lblVisite11.DoubleClick
        majVisite(entree, sortie, tabEPD(indexEPD), lblCoup11.Text, lblVisite11.Text)
    End Sub

    Private Sub lblVisite12_DoubleClick(sender As Object, e As EventArgs) Handles lblVisite12.DoubleClick
        majVisite(entree, sortie, tabEPD(indexEPD), lblCoup12.Text, lblVisite12.Text)
    End Sub

    Private Sub lblVisite13_DoubleClick(sender As Object, e As EventArgs) Handles lblVisite13.DoubleClick
        majVisite(entree, sortie, tabEPD(indexEPD), lblCoup13.Text, lblVisite13.Text)
    End Sub

    Private Sub lblVisite14_DoubleClick(sender As Object, e As EventArgs) Handles lblVisite14.DoubleClick
        majVisite(entree, sortie, tabEPD(indexEPD), lblCoup14.Text, lblVisite14.Text)
    End Sub

    Private Sub lblVisite15_DoubleClick(sender As Object, e As EventArgs) Handles lblVisite15.DoubleClick
        majVisite(entree, sortie, tabEPD(indexEPD), lblCoup15.Text, lblVisite15.Text)
    End Sub

    Private Sub lblVisite16_DoubleClick(sender As Object, e As EventArgs) Handles lblVisite16.DoubleClick
        majVisite(entree, sortie, tabEPD(indexEPD), lblCoup16.Text, lblVisite16.Text)
    End Sub

    Private Sub lblVisite17_DoubleClick(sender As Object, e As EventArgs) Handles lblVisite17.DoubleClick
        majVisite(entree, sortie, tabEPD(indexEPD), lblCoup17.Text, lblVisite17.Text)
    End Sub

    Private Sub lblVisite18_DoubleClick(sender As Object, e As EventArgs) Handles lblVisite18.DoubleClick
        majVisite(entree, sortie, tabEPD(indexEPD), lblCoup18.Text, lblVisite18.Text)
    End Sub

    Private Sub lblVisite19_DoubleClick(sender As Object, e As EventArgs) Handles lblVisite19.DoubleClick
        majVisite(entree, sortie, tabEPD(indexEPD), lblCoup19.Text, lblVisite19.Text)
    End Sub

    Private Sub lblVisite20_DoubleClick(sender As Object, e As EventArgs) Handles lblVisite20.DoubleClick
        majVisite(entree, sortie, tabEPD(indexEPD), lblCoup20.Text, lblVisite20.Text)
    End Sub

    Private Sub majVisite(entreeVisite As System.IO.StreamWriter, sortieVisite As System.IO.StreamReader, epd As String, coup As String, visite As String)
        Dim critereEXP(23) As Byte, i As Long, nouvelleVisite As Integer, reponse As String, visiteActuelle As Integer

        If coup = "" Then
            Exit Sub
        End If
        Cursor = Cursors.WaitCursor
        txtRecherche.Enabled = False

        visiteActuelle = CInt(Replace(visite, "x", ""))

        'position                coup   ?   ?    score             profondeur        visites
        '0 1  2  3  4  5  6  7   8  9   a   b    c   d   e   f     0   1   2   3     4   5   6   7
        '0 +1 +2 +3 +4 +5 +6 +7  +8 +9  +10 +11  +12 +13 +14 +15   +16 +17 +18 +19   +20 +21 +22 +23
        'tabExperience(i+16) = 60 (profondeur)
        'tabExperience(i+20) = 120 (visites)

        reponse = InputBox("New count value ?", "min 0 <= count value <= max 255", visiteActuelle)
        If reponse <> "" Then
            If IsNumeric(reponse) Then
                nouvelleVisite = CInt(reponse)
                If 0 <= nouvelleVisite And nouvelleVisite <= 255 And nouvelleVisite <> visiteActuelle Then
                    'on prépare un critère de recherche basé sur la position en cours et le coup sélectionné
                    Array.Copy(epdToEXP(entreeVisite, sortieVisite, epd), 0, critereEXP, 0, 8)
                    Array.Copy(moveToEXP(coup), 0, critereEXP, 8, 2)

                    'on décharge le moteur
                    dechargerMoteur()
                    Application.DoEvents()

                    'on met à jour le compteur de visites dans Eman.exp
                    '0-25 = SugaR Experience version 2
                    For i = 26 To UBound(tabExperience) Step 24
                        'position
                        If tabExperience(i) = critereEXP(0) Then
                            If tabExperience(i + 1) = critereEXP(1) Then
                                If tabExperience(i + 2) = critereEXP(2) Then
                                    If tabExperience(i + 3) = critereEXP(3) Then
                                        If tabExperience(i + 4) = critereEXP(4) Then
                                            If tabExperience(i + 5) = critereEXP(5) Then
                                                If tabExperience(i + 6) = critereEXP(6) Then
                                                    If tabExperience(i + 7) = critereEXP(7) Then
                                                        'coup
                                                        If tabExperience(i + 8) = critereEXP(8) And tabExperience(i + 9) = critereEXP(9) Then
                                                            'visite
                                                            If tabExperience(i + 20) <> nouvelleVisite Then
                                                                tabExperience(i + 20) = nouvelleVisite
                                                                My.Computer.FileSystem.WriteAllBytes(lblExperience.Text, tabExperience, False)
                                                                'dans tous les cas on sort ! (Sinon si le fichier n'était pas défragmenté, ça peut boucler très très longtemps)
                                                                Exit For
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next

                    'on recharge le moteur
                    chargerMoteur(lblEman.Text, lblExperience.Text)

                    'on met à jour l'affichage
                    If indexEPD >= 0 And moteur_court <> "" Then
                        If tabEPD(indexEPD) <> "" Then
                            afficherEXP(tabEPD(indexEPD))
                        End If
                    End If

                End If
            End If
        End If

        txtRecherche.Enabled = True
        Cursor = Cursors.Default
    End Sub

    Private Sub lblInfo_MouseClick(sender As Object, e As MouseEventArgs) Handles lblInfo.MouseClick
        If e.Button = Windows.Forms.MouseButtons.Right Then
            If lblRecherche.Text = "Position EPD :" Then
                If MsgBox("Copy " & tabEPD(indexEPD) & " into clipboard ?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    My.Computer.Clipboard.SetText(tabEPD(indexEPD))
                End If
            ElseIf lblRecherche.Text = "Suite UCI :" Then
                If MsgBox("Copy " & tabEPD(indexEPD) & " into clipboard ?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    My.Computer.Clipboard.SetText(tabEPD(indexEPD))
                ElseIf MsgBox("Copy " & txtRecherche.Text & " into clipboard ?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    My.Computer.Clipboard.SetText(txtRecherche.Text)
                End If
            ElseIf lblRecherche.Text = "Mode PGN :" Then
                If MsgBox("Copy " & lblInfo.Text.Substring(lblInfo.Text.IndexOf(":") + 2) & " into clipboard ?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    My.Computer.Clipboard.SetText(lblInfo.Text.Substring(lblInfo.Text.IndexOf(":") + 2))
                End If
            End If
        End If

    End Sub

End Class

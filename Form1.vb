Imports System.Data
Imports Microsoft.Data.Sqlite
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports System.IO
Imports System.Reflection.Metadata
Imports Document = iTextSharp.text.Document
Imports System.Data.OleDb
Imports System.Diagnostics.Eventing.Reader
Imports System.Diagnostics.Metrics
Imports System.ComponentModel
Imports System.Net.Mail

Public Class PdfPageEvent

    Inherits PdfPageEventHelper
    Private ReadOnly _titre As String
    Private ReadOnly _titrefont As iTextSharp.text.Font
    Private ReadOnly img As Image
    Private ReadOnly numProd As String
    Private ReadOnly typeVI As String
    Private poste As String = ""
    Private bddLocal As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Setup\BDD\01_Configuration\Config_Bdd.mdb;Jet OLEDB:Database Password=ACMAT"
    Private queryLocal As String = "SELECT poste FROM 00_07b_Ligne_de_production WHERE Actif = 'OUI'"
    Dim accessConnection As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\vcn.ds.volvo.net\cli-sd\sd0627\047962\03_SUIVI_PROD\02_LIGNE_DE_PRODUCTION_VBL\00_Template\VBL_EOP_Template.mdb")

    Public Sub New(image As Image, numProd As String, typeVI As String)
        img = image

        Me.numProd = numProd
        Me.typeVI = typeVI
        Me.poste = poste
    End Sub
    Public Sub New(titre As String, titrefont As iTextSharp.text.Font)
        _titre = titre
        _titrefont = titrefont
    End Sub

    Public Overrides Sub OnEndPage(writer As PdfWriter, document As Document)
        Using connection As New OleDbConnection(bddLocal)
            Using command As New OleDbCommand(queryLocal, connection)
                connection.Open()
                Dim reader As OleDbDataReader = command.ExecuteReader()
                If reader.Read() Then
                    poste = reader.GetString(0)
                End If
                reader.Close()
            End Using
            connection.Close()
        End Using
        MyBase.OnEndPage(writer, document)
        img.SetAbsolutePosition(-5, 505)
        img.ScaleToFit(100, 100)
        document.Add(img)
        Dim titrefont = FontFactory.GetFont("Arial", 28, Font.BOLD, BaseColor.BLACK)
        Dim titreText As String = "Rapport Serrage " & poste & " : " & typeVI & " " & numProd

        Dim titreParagraph As New Paragraph(_titre, _titrefont)
        titreParagraph.Alignment = Element.ALIGN_CENTER


        Dim cb As PdfContentByte = writer.DirectContent


        ColumnText.ShowTextAligned(cb, Element.ALIGN_CENTER, New Phrase(titreText, titrefont),
                               document.PageSize.Width / 2, document.PageSize.Height - 50, 0)


        ColumnText.ShowTextAligned(cb, Element.ALIGN_CENTER, New Phrase(" "),
                               document.PageSize.Width / 2, document.PageSize.Height - 65, 0)
        ColumnText.ShowTextAligned(cb, Element.ALIGN_CENTER, New Phrase(" "),
                               document.PageSize.Width / 2, document.PageSize.Height - 80, 0)
    End Sub

End Class

Public Class Form1
    Dim typeVI As String = ""
    Dim numProd As String = ""
    Dim fichierVI As String = ""
    Dim poste As String = ""
    Dim bddLocal As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Setup\BDD\01_Configuration\Config_Bdd.mdb;Jet OLEDB:Database Password=ACMAT"
    Dim queryLocal As String = "SELECT poste FROM 00_07b_Ligne_de_production WHERE Actif = 'OUI'"
    Dim bddChargement As String = ""
    Dim queryTypeVI As String = "SELECT Modèle_véhicule FROM 01_01_EOP_POSTE_RESEAU WHERE numero_de_production = '" & numProd & "'"
    Public Function GetTotalAttendu(idSerrage As String) As Integer

        Dim accessConnection As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\vcn.ds.volvo.net\cli-sd\sd0627\047962\03_SUIVI_PROD\02_LIGNE_DE_PRODUCTION_VBL\00_Template\VBL_EOP_Template.mdb")
        accessConnection.Open()
        Dim cmd As New OleDbCommand("SELECT Nombre FROM 03_01_Tbl_Serrages_Config WHERE ID_Serrage = @id", accessConnection)
        cmd.Parameters.AddWithValue("@id", idSerrage)
        Dim result As Object = cmd.ExecuteScalar()
        Return If(result IsNot Nothing, Convert.ToInt32(result), 0)
    End Function
    Public Function GetSerragesNonConformes() As List(Of String)

        Dim serragesNonConformes As New List(Of String)


        serragesNonConformes.Add("Serrage A - Non conforme")
        serragesNonConformes.Add("Serrage B - Non conforme")

        Return serragesNonConformes
    End Function
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Using connection As New OleDbConnection(bddLocal)
            Using command As New OleDbCommand(queryLocal, connection)
                connection.Open()
                Dim reader As OleDbDataReader = command.ExecuteReader()
                If reader.Read() Then
                    poste = reader.GetString(0)
                End If
                reader.Close()
            End Using
            connection.Close()
        End Using

        Dim queryChargement As String = "SELECT numero_de_production,Nom_du_fichier FROM 01_01_EOP_POSTE_RESEAU WHERE Poste_de_production = '" & poste & "'"

        Select Case True
            Case poste.StartsWith("VB")
                bddChargement = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\vcn.ds.volvo.net\cli-sd\sd0627\047962\03_SUIVI_PROD\02_LIGNE_DE_PRODUCTION_VBL\02_Base_données_Chargement_MV\Chargement_VBL.mdb"
            Case poste.StartsWith("R")
                bddChargement = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\vcn.ds.volvo.net\cli-sd\sd0627\047962\03_SUIVI_PROD\04_LIGNE_DE_PRODUCTION_GBC-VLRA\02_Base_données_Chargement_MV\Chargement_GBC.mdb"
            Case poste.StartsWith("T")
                bddChargement = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\vcn.ds.volvo.net\cli-sd\sd0627\047962\03_SUIVI_PROD\03_LIGNE_DE_PRODUCTION_TRM10000\02_Base_données_Chargement_MV\Chargement_TRM10000.mdb"
            Case poste.StartsWith("PV")
                bddChargement = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\vcn.ds.volvo.net\cli-sd\sd0627\047962\03_SUIVI_PROD\05_LIGNE_DE_PRODUCTION_PVP\02_Base_données_Chargement_MV\Chargement_PVP.mdb"
            Case Else
                MsgBox("erreur!")
        End Select
        Using connection As New OleDbConnection(bddChargement)
            Using command As New OleDbCommand(queryChargement, connection)
                connection.Open()
                Dim reader As OleDbDataReader = command.ExecuteReader()
                If reader.Read() Then
                    numProd = reader.GetString(0)
                    fichierVI = reader.GetString(1)
                End If
                reader.Close()
            End Using
            connection.Close()
        End Using

        Dim queryTypeVI As String = "SELECT Modèle_véhicule FROM 01_01_EOP_POSTE_RESEAU WHERE numero_de_production = '" & numProd & "'"

        Using connection As New OleDbConnection(bddChargement)
            Using command As New OleDbCommand(queryTypeVI, connection)
                connection.Open()
                Dim reader As OleDbDataReader = command.ExecuteReader()
                If reader.Read() Then
                    ' typeVI = reader.GetString(0)
                End If
                reader.Close()
            End Using
            connection.Close()
        End Using
        typeVI = "VBL"
        Dim sourceDBPath As String = "\\vcn.ds.volvo.net\cli-sd\sd0627\047962\07_HARDWARE\SERRAGE\02_LIGNE_DE_PRODUCTION_VBL\TEST\VB40-1682-4-40.db"
        Dim targetDBPath As String = "\\vcn.ds.volvo.net\cli-sd\sd0627\047962\07_HARDWARE\SERRAGE\02_LIGNE_DE_PRODUCTION_VBL\TEST"
        If poste = "VB10" Then
            If Not String.IsNullOrEmpty(numProd) Then
                targetDBPath = "C:\Users\A479250\Desktop\db csv\" & numProd & ".db"
                File.Copy(sourceDBPath, targetDBPath, True)

            Else

            End If
        Else
            targetDBPath = "C:\Users\A479250\Desktop\db csv\" & numProd & ".db"
            Dim connectionString As String = "Data Source=" & targetDBPath

            Try
                Using sourceConn As New SqliteConnection("Data Source=" & sourceDBPath)
                    sourceConn.Open()

                    Using targetConn As New SqliteConnection(connectionString)
                        targetConn.Open()



                    End Using
                End Using
            Catch ex As Exception

            End Try
        End If

    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim sqlite_conn As SqliteConnection = Nothing
        Dim access_conn As OleDbConnection = Nothing
        Dim myCmd As SqliteCommand = Nothing
        Dim access_cmd As OleDbCommand = Nothing
        Try

            sqlite_conn = New SqliteConnection("Data Source=\\vcn.ds.volvo.net\cli-sd\sd0627\047962\07_HARDWARE\SERRAGE\02_LIGNE_DE_PRODUCTION_VBL\TEST\VB40-1682-4-40.db")
            sqlite_conn.Open()
            myCmd = sqlite_conn.CreateCommand()
            myCmd.CommandText = "SELECT jd.id, jd.time, js.name, js.torque, jd.max_torque, jd.serno FROM joint_data jd INNER JOIN joint_setting js ON jd.settings_id = js.id"
            Dim myReader As SqliteDataReader = myCmd.ExecuteReader()
            access_conn = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\vcn.ds.volvo.net\cli-sd\sd0627\047962\03_SUIVI_PROD\02_LIGNE_DE_PRODUCTION_VBL\01_Base_données_VI\01_Ligne_de_production\" & typeVI & "_EOP_" & numProd & ".mdb")
            access_conn.Open()
            access_cmd = access_conn.CreateCommand()

            Dim i As Integer = 0
            While myReader.Read()
                Dim unixTime As Long = Convert.ToInt64(myReader("time"))
                Dim dateTime As DateTime = New DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc).AddSeconds(unixTime)
                Dim name As String = myReader("name").ToString()
                Dim torque As Double = Convert.ToDouble(myReader("torque").ToString())
                Dim max_torque As Double = Convert.ToDouble(myReader("max_torque")) / 1000.0
                Dim serno As String = myReader("serno").ToString()
                Dim codeUnique As String = typeVI & "-" & poste & "-" & numProd

                Dim difference As Double = Math.Abs(max_torque - torque)
                Dim tolerance As Double = 20
                Dim conforme As String = "Conforme"
                If name.StartsWith("PS") Then
                    tolerance = 10
                End If
                If difference > max_torque * (tolerance / 100.0) Then
                    conforme = "Non Conforme"
                End If

                Dim max_torqueFormatted As String = max_torque.ToString("F3")
                Dim torqueFormatted As String = torque.ToString("F3")


                access_cmd.CommandText = "INSERT INTO 03_02_Tbl_Serrages_ALL (code_unique, date_serrage, Nom, valeur_cible, serrage_realise, serno, resultat, poste) VALUES (@id, @time, @name, @torque, @max_torque, @serno, @resultat, @poste)"
                access_cmd.Parameters.Clear()
                access_cmd.Parameters.AddWithValue("@id", codeUnique & "-" & i)
                access_cmd.Parameters.AddWithValue("@time", dateTime)
                access_cmd.Parameters.AddWithValue("@name", name)
                access_cmd.Parameters.AddWithValue("@torque", torqueFormatted)
                access_cmd.Parameters.AddWithValue("@max_torque", max_torqueFormatted)
                access_cmd.Parameters.AddWithValue("@serno", serno)
                access_cmd.Parameters.AddWithValue("@resultat", conforme)
                access_cmd.Parameters.AddWithValue("@poste", poste)
                access_cmd.ExecuteNonQuery()
                i = i + 1
            End While


            MessageBox.Show("Insertion terminée avec succès.")

        Catch ex As Exception
            MessageBox.Show("Une erreur est survenue : " & ex.Message)
        Finally

            If myCmd IsNot Nothing Then myCmd.Dispose()
            If access_cmd IsNot Nothing Then access_cmd.Dispose()
            If sqlite_conn IsNot Nothing AndAlso sqlite_conn.State = ConnectionState.Open Then sqlite_conn.Close()
            If access_conn IsNot Nothing AndAlso access_conn.State = ConnectionState.Open Then access_conn.Close()
        End Try
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim databasePath As String = "\\vcn.ds.volvo.net\cli-sd\sd0627\047962\03_SUIVI_PROD\02_LIGNE_DE_PRODUCTION_VBL\01_Base_données_VI\01_Ligne_de_production\" & typeVI & "_EOP_" & numProd & ".mdb"
        Dim accessConn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & databasePath)
        Try

            accessConn.Open()

        Catch ex As OleDb.OleDbException
            MessageBox.Show("Erreur de connexion : " & ex.Message)
        Catch ex As Exception
            MessageBox.Show("Erreur générale : " & ex.Message)

        End Try



        Dim accessCmd As OleDbCommand
        Dim accessReader As OleDbDataReader
        Dim rowsTable As New List(Of List(Of String))()
        Dim rowsSummaryPS As New List(Of List(Of String))()
        Dim rowsSummaryVB As New List(Of List(Of String))()
        Dim test As Integer = 0


        accessCmd = New OleDbCommand("SELECT code_unique, date_serrage, Nom, valeur_cible, serrage_realise, serno, poste FROM 03_02_Tbl_Serrages_ALL", accessConn)
        accessReader = accessCmd.ExecuteReader()

        Dim document As New Document(PageSize.A4.Rotate(), 36, 36, 100, 36)
        Dim output As New FileStream("C:\Users\A479250\Desktop\VB40-1682-4-40.pdf", FileMode.Create)
        Dim writer = PdfWriter.GetInstance(document, output)
        document.Open()

        Dim font As iTextSharp.text.Font = FontFactory.GetFont(FontFactory.HELVETICA, 14)
        Dim fontEntete As iTextSharp.text.Font = FontFactory.GetFont(FontFactory.HELVETICA, 17)

        Dim pageWidth As Single = document.PageSize.Width - document.LeftMargin - document.RightMargin
        Dim image As iTextSharp.text.Image = iTextSharp.text.Image.GetInstance(vraiPdfExport.My.Resources.Arquus, System.Drawing.Imaging.ImageFormat.Png)
        Dim pageEvent As New PdfPageEvent(image, numProd, typeVI)

        ' Dim titre As New Paragraph(titreText, titrefont)
        ' titre.Alignment = Element.ALIGN_CENTER

        document.Open()
        writer.PageEvent = pageEvent
        '  document.Add(titre)

        Dim pdfTable As New PdfPTable(4)
        pdfTable.TotalWidth = pageWidth
        pdfTable.LockedWidth = True
        pdfTable.SplitLate = False
        pdfTable.AddCell(New PdfPCell(New Phrase("Nom", fontEntete)) With {.BackgroundColor = New BaseColor(173, 216, 230), .HorizontalAlignment = Element.ALIGN_CENTER})
        pdfTable.AddCell(New PdfPCell(New Phrase("Conforme ?", fontEntete)) With {.BackgroundColor = New BaseColor(173, 216, 230), .HorizontalAlignment = Element.ALIGN_CENTER})
        pdfTable.AddCell(New PdfPCell(New Phrase("Date", fontEntete)) With {.BackgroundColor = New BaseColor(173, 216, 230), .HorizontalAlignment = Element.ALIGN_CENTER})
        pdfTable.AddCell(New PdfPCell(New Phrase("Serrage réalisé", fontEntete)) With {.BackgroundColor = New BaseColor(173, 216, 230), .HorizontalAlignment = Element.ALIGN_CENTER})

        Dim pages As Integer = 0
        Dim totalConformes As Integer = 0
        Dim totalNonConformes As Integer = 0
        Dim summaryDataVB As New Dictionary(Of String, Dictionary(Of String, Integer))()
        Dim summaryDataPS As New Dictionary(Of String, Dictionary(Of String, Integer))()
        Dim rows As New List(Of (String, DateTime, String, Double, Double, String, String, String))()


        While accessReader.Read()
            Dim name As String = accessReader("Nom").ToString()
            Dim id As String = accessReader("code_unique").ToString()
            Dim dateTime As DateTime = Convert.ToDateTime(accessReader("date_serrage"))
            Dim torque As Double = Convert.ToDouble(accessReader("valeur_cible"))
            Dim maxTorque As Double = Convert.ToDouble(accessReader("serrage_realise"))
            Dim serno As String = accessReader("serno").ToString()
            Dim posteCmd As OleDbCommand = New OleDbCommand("SELECT poste FROM 03_02_Tbl_Serrages_ALL WHERE nom = @nom", accessConn)
            posteCmd.Parameters.AddWithValue("@nom", name)
            Dim poste As String = Convert.ToString(posteCmd.ExecuteScalar())

            If String.IsNullOrEmpty(poste) Then
                poste = "Non défini"
            End If

            Dim conforme As String = "Conforme"
            Dim difference As Double = Math.Abs(maxTorque - torque)
            Dim tolerance As Double


            If name.StartsWith("VB") Then
                tolerance = 20
            Else
                tolerance = 10
            End If

            Dim tolerancePercentage As Double = tolerance / 100.0


            If difference > maxTorque * tolerancePercentage Then
                conforme = "Non Conforme"
            Else
                conforme = "Conforme"
            End If

            rows.Add((id, dateTime, name, torque, maxTorque, serno, conforme, poste))
        End While

        Dim indiceAEnlever As New List(Of Integer)()

        For i As Integer = 0 To rows.Count - 1
            Dim row = rows(i)
            Dim idCell As String = row.Item1
            Dim timeCell As String = row.Item2.ToString("yyyy-MM-dd HH:mm:ss")
            Dim nameCell As String = row.Item3
            Dim torqueCell As String = row.Item4.ToString("F3")
            Dim maxTorqueCell As String = row.Item5.ToString("F3")
            Dim sernoCell As String = row.Item6
            Dim resultCell As String = row.Item7
            Dim posteCell As String = row.Item8

            Dim nameLigne3 As String = rows(i).Item3
            Dim conformeLigne3 As String = rows(i).Item7




            If resultCell = "Non Conforme" Then

                If i < rows.Count - 1 AndAlso rows(i + 1).Item5 < 0 Then
                    For j As Integer = 0 To i - 1
                        If rows(j).Item3 = nameLigne3 Then
                            If conformeLigne3 = "Conforme" Then

                            End If
                        End If
                    Next
                    indiceAEnlever.Add(i)
                    indiceAEnlever.Add(i + 1)

                    i += 1
                    Continue For
                Else
                    test = test + 1
                    Dim access_conn As OleDbConnection = Nothing

                    Dim access_cmd As OleDbCommand = Nothing


                    Try

                        Dim listeNonConforme As New List(Of List(Of String))()
                        access_conn = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & databasePath)
                        access_conn.Open()

                        access_cmd = access_conn.CreateCommand()
                        access_cmd.CommandText = "INSERT INTO 03_03_Tbl_Non_Conformes_ALL (code_unique, date_serrage, Nom, valeur_cible, serrage_realise, serno, resultat, poste) VALUES (@id, @time, @name, @torque, @max_torque, @serno, @resultat, @poste)"

                        access_cmd.Parameters.Clear()
                        access_cmd.Parameters.AddWithValue("@id", idCell)
                        access_cmd.Parameters.AddWithValue("@time", timeCell)
                        access_cmd.Parameters.AddWithValue("@name", nameCell)
                        access_cmd.Parameters.AddWithValue("@torque", torqueCell)
                        access_cmd.Parameters.AddWithValue("@max_torque", maxTorqueCell)
                        access_cmd.Parameters.AddWithValue("@serno", sernoCell)
                        access_cmd.Parameters.AddWithValue("@resultat", resultCell)
                        access_cmd.Parameters.AddWithValue("@poste", posteCell)


                        access_cmd.ExecuteNonQuery()


                        Dim SmtpServer As New SmtpClient("mailgot.it.volvo.net", 25)
                        SmtpServer.Credentials = New Net.NetworkCredential("industrial.id.A361145@VOLVO.com", "")
                        SmtpServer.EnableSsl = False


                        Dim mail As New MailMessage()
                        mail.From = New MailAddress("industrial.id.A361145@VOLVO.com")
                        mail.Subject = "Serrage non conforme- " & nameCell


                        mail.Body = "Nom : " & nameCell & vbCrLf &
                                    "Code unique : " & idCell & vbCrLf &
                                    "Date de serrage : " & timeCell & vbCrLf &
                                    "Serrage réalisé : " & maxTorqueCell & vbCrLf &
                                    "Poste : " & posteCell & vbCrLf &
                                    "Résultat : " & resultCell


                        ' mail.To.Add("arthur.bompoil@arquus-defense.com")
                        ' SmtpServer.Send(mail)
                    Finally
                        If access_conn IsNot Nothing Then
                            access_conn.Close()
                        End If
                    End Try
                End If
            End If




        Next

        For Each index In indiceAEnlever.Distinct().OrderByDescending(Function(i) i)
            rows.RemoveAt(index)
        Next


        For Each row In rows


            Dim nameCell As String = row.Item3
            Dim resultCell As String = row.Item7

            If nameCell.StartsWith("PS") Then

                If Not summaryDataPS.ContainsKey(nameCell) Then
                    summaryDataPS(nameCell) = New Dictionary(Of String, Integer)()
                    summaryDataPS(nameCell)("Conformes") = 0
                    summaryDataPS(nameCell)("Non conformes") = 0
                End If


                If resultCell = "Conforme" Then
                    summaryDataPS(nameCell)("Conformes") += 1
                Else
                    summaryDataPS(nameCell)("Non conformes") += 1
                End If

            ElseIf nameCell.StartsWith("VB") Then

                If Not summaryDataVB.ContainsKey(nameCell) Then
                    summaryDataVB(nameCell) = New Dictionary(Of String, Integer)()
                    summaryDataVB(nameCell)("Conformes") = 0
                    summaryDataVB(nameCell)("Non conformes") = 0
                End If


                If resultCell = "Conforme" Then
                    summaryDataVB(nameCell)("Conformes") += 1
                Else
                    summaryDataVB(nameCell)("Non conformes") += 1
                End If
            End If

            If nameCell.StartsWith("VB") OrElse resultCell <> "Conforme" Then
                Continue For
            End If

            pdfTable.AddCell(row.Item3)
            pdfTable.AddCell(row.Item7)
            pdfTable.AddCell(row.Item2.ToString("yyyy-MM-dd HH:mm:ss"))
            pdfTable.AddCell(row.Item5.ToString("F3"))

            pages += 1
            If pages = 27 Then
                ' document.Add(pdfTable)
                'document.NewPage()
                pages = 0
                ' pdfTable.Rows.Clear()
                'document.Add(titre)

                pdfTable.AddCell(New PdfPCell(New Phrase("Nom", fontEntete)) With {.BackgroundColor = New BaseColor(173, 216, 230), .HorizontalAlignment = Element.ALIGN_CENTER})
                pdfTable.AddCell(New PdfPCell(New Phrase("Conforme ?", fontEntete)) With {.BackgroundColor = New BaseColor(173, 216, 230), .HorizontalAlignment = Element.ALIGN_CENTER})
                pdfTable.AddCell(New PdfPCell(New Phrase("Date", fontEntete)) With {.BackgroundColor = New BaseColor(173, 216, 230), .HorizontalAlignment = Element.ALIGN_CENTER})
                pdfTable.AddCell(New PdfPCell(New Phrase("Serrage réalisé", fontEntete)) With {.BackgroundColor = New BaseColor(173, 216, 230), .HorizontalAlignment = Element.ALIGN_CENTER})
            End If

        Next


        Dim fontrecap As iTextSharp.text.Font = FontFactory.GetFont(FontFactory.HELVETICA, 15)

        Dim pdfTableSummaryPS As New PdfPTable(3)
        pdfTableSummaryPS.TotalWidth = pageWidth
        pdfTableSummaryPS.LockedWidth = True
        pdfTableSummaryPS.AddCell(New PdfPCell(New Phrase("Nom", fontrecap)) With {.BackgroundColor = New BaseColor(173, 216, 230), .HorizontalAlignment = Element.ALIGN_CENTER})
        pdfTableSummaryPS.AddCell(New PdfPCell(New Phrase("Total conformes", fontrecap)) With {.BackgroundColor = New BaseColor(173, 216, 230), .HorizontalAlignment = Element.ALIGN_CENTER})
        pdfTableSummaryPS.AddCell(New PdfPCell(New Phrase("Total attendu", fontrecap)) With {.BackgroundColor = New BaseColor(173, 216, 230), .HorizontalAlignment = Element.ALIGN_CENTER})

        Dim sortedSummaryPS = summaryDataPS.OrderBy(Function(entry) entry.Key).ToList()


        Dim accessConnection As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\vcn.ds.volvo.net\cli-sd\sd0627\047962\03_SUIVI_PROD\02_LIGNE_DE_PRODUCTION_VBL\00_Template\VBL_EOP_Template.mdb")
        accessConnection.Open()

        For Each entry In sortedSummaryPS
            Dim totalAttendu As Integer = GetTotalAttendu(entry.Key)
            pdfTableSummaryPS.AddCell(New PdfPCell(New Phrase(entry.Key, fontrecap)) With {.HorizontalAlignment = Element.ALIGN_CENTER})
            pdfTableSummaryPS.AddCell(New PdfPCell(New Phrase(entry.Value("Conformes").ToString(), fontrecap)) With {.HorizontalAlignment = Element.ALIGN_CENTER})
            pdfTableSummaryPS.AddCell(New PdfPCell(New Phrase(totalAttendu.ToString(), fontrecap)) With {.HorizontalAlignment = Element.ALIGN_CENTER})
        Next


        document.Add(pdfTableSummaryPS)
        document.NewPage()



        Dim pdfTableSummaryVB As New PdfPTable(3)
        pdfTableSummaryVB.TotalWidth = pageWidth
        pdfTableSummaryVB.LockedWidth = True
        pdfTableSummaryVB.AddCell(New PdfPCell(New Phrase("Nom", fontrecap)) With {.BackgroundColor = New BaseColor(173, 216, 230), .HorizontalAlignment = Element.ALIGN_CENTER})
        pdfTableSummaryVB.AddCell(New PdfPCell(New Phrase("Total conformes", fontrecap)) With {.BackgroundColor = New BaseColor(173, 216, 230), .HorizontalAlignment = Element.ALIGN_CENTER})
        pdfTableSummaryVB.AddCell(New PdfPCell(New Phrase("Total attendu", fontrecap)) With {.BackgroundColor = New BaseColor(173, 216, 230), .HorizontalAlignment = Element.ALIGN_CENTER})

        Dim sortedSummaryVB = summaryDataVB.OrderBy(Function(entry) entry.Key).ToList()


        For Each entry In sortedSummaryVB
            Dim totalAttendu As Integer = GetTotalAttendu(entry.Key)
            pdfTableSummaryVB.AddCell(New PdfPCell(New Phrase(entry.Key, fontrecap)) With {.HorizontalAlignment = Element.ALIGN_CENTER})
            pdfTableSummaryVB.AddCell(New PdfPCell(New Phrase(entry.Value("Conformes").ToString(), fontrecap)) With {.HorizontalAlignment = Element.ALIGN_CENTER})
            pdfTableSummaryVB.AddCell(New PdfPCell(New Phrase(totalAttendu.ToString(), fontrecap)) With {.HorizontalAlignment = Element.ALIGN_CENTER})
        Next

        document.Add(pdfTableSummaryVB)
        document.NewPage()

        accessConnection.Close()
        '  document.Add(titre)

        document.Add(pdfTable)


        document.Close()
        writer.Close()
        output.Close()
        accessReader.Close()
        accessConn.Close()


    End Sub


End Class

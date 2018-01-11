'Option Explicit On
Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.OleDb.OleDbConnection
Imports Microsoft.Office.Interop
Imports Microsoft.VisualBasic
Imports System.Data.SqlClient

'1111111111111111111111111111111111111
'1111111111111111111111111111111111111

'Imports excel = Microsoft.Office.Interop.Excel

'Public Class ClasseExcel
'    Private objexcel As New excel.Application
'    Dim xlBook As excel.Workbook
'    Dim xlworksheet As excel.Worksheet
'    Public Sub New()
'        xlBook = objexcel.Workbooks.Add
'        xlworksheet = CType(xlBook.ActiveSheet, excel.Worksheet)
'    End Sub

'    Public Sub Writelist(ByVal mylist As List(Of String))
'        ' écrit dans la colonne A1 :: A?  mylist
'        Dim cellstrcopy As String
'        Dim indexcol As Integer
'        Dim indexrow As Integer
'        Dim myfont As Font
'        myfont = New Font("arial", 12, FontStyle.Bold)
'        cellstrcopy = String.Empty
'        Try
'            With xlworksheet
'                indexcol = 1
'                indexrow = 1
'                cellstrcopy = Convert.ToChar(indexcol + 64) & indexrow.ToString
'                For Each item In mylist

'                    .Cells(indexrow, indexcol) = item

'                    .Range(cellstrcopy).BorderAround()
'                    .Range(cellstrcopy).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LemonChiffon)
'                    .Range(cellstrcopy).Select()
'                    .Range(cellstrcopy).HorizontalAlignment = excel.XlVAlign.xlVAlignCenter
'                    With .Range(cellstrcopy).Font
'                        .Name = "Arial"
'                        .Strikethrough = False
'                        .Bold = True
'                        .Size = 12
'                    End With
'                    indexrow += 1
'                    cellstrcopy = Convert.ToChar(indexcol + 64) & indexrow.ToString
'                Next

'            End With
'            objexcel.Visible = True
'            objexcel = Nothing
'        Catch ex As Exception
'            MessageBox.Show(ex.Message.ToString)
'        End Try

'    End Sub
'End Class

'1111111111111111111111111111111111111
'1111111111111111111111111111111111111

Public Class Form1

    'Dim xlAppSource As Excel.Application
    'Dim xlWorkBookSource As Excel.Workbook
    'Dim xlWorkSheetSource As Excel.Worksheet
    'Dim xlSourceRange As Excel.Range

    'Dim xlAppModel As Excel.Application
    'Dim xlWorkBookModel As Excel.Workbook
    'Dim xlWorkSheetModel As Excel.Worksheet
    'Dim xlModelRange As Excel.Range

    Dim xlAppRapport As Excel.Application
    Dim xlWorkBookRapport As Excel.Workbook
    Dim xlWorkSheetRapport As Excel.Worksheet
    Dim xlRapportRange As Excel.Range

    Dim RepportPath As String

    Dim Affichage As String

    Dim NombreSup As Integer

    Dim sFileName As String = ""

    Dim sFileI As String = ""

    Dim sDirectory As String = "C:\Program RIC\Adolphe Expert\RIC\Source"
    Dim mDirectory As String = "C:\Program RIC\Adolphe Expert\RIC\Model"
    Dim riDirectory As String = "C:\Program RIC\Adolphe Expert\RIC\Rapport Initial"
    Dim rfDirectory As String = "C:\Program RIC\Adolphe Expert\RIC\Rapport Final"

    'Dim sFileEducation As String = "D:\Source\Clone de EDUCATION - labels - 2017-12-22-08-36.xlsx"
    Dim sFileEducation As String = "C:\Program RIC\Adolphe Expert\RIC\Source\Clone de EDUCATION - labels - 2017-12-22-08-36.xlsx"

    'Dim mFileEducation As String = "D:\Model\EDUCATION  21.12.17 V2.xlsx"
    'Dim mFileEducation As String = "C:\Program RIC\Adolphe Expert\RIC\Model\EDUCATION  21.12.17 V2.xlsx"
    Dim mFileEducation As String = "C:\Program RIC\Adolphe Expert\RIC\Model\EDUCATION  09.01.2018.xltx"
    '‪C:\Program RIC\Adolphe Expert\RIC\Model\EDUCATION  21.12.17 V2.xltx

    'Dim path As String = "D:\EDUCATION - labels - 2017-11-22-14-22.xlsx"

    Dim nRapportEducation As String = "Enquête"
    Dim nDistrictEducation As String = "Disctrict"

    'Affichage = "Null"

    'Dim Truc As CheckBox

    Public Sub NAR(ByVal o As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(o)
        Catch ex As Exception
        Finally
            o = Nothing
        End Try
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'yes yes
        'Dim TreeViewFiltresGeographiques As TreeView
        'TreeViewFiltresGeographiques = New TreeView()
        'Me.Controls.Add(TreeViewFiltresGeographiques)
        'TreeViewFiltresGeographiques.Nodes.Clear()
        'TreeViewFiltresGeographiques.SelectedNode.Checked = True

        'For Each Truc In GroupBoxChoixInfras.Controls
        '    If Truc.Checked = True Then
        '        ComboBoxInfraAFiltrer.Text = Truc.Text
        '    End If
        'Next Truc

        GroupBoxAdministration.Hide()
        GroupBoxEauEtAssainissement.Hide()
        GroupBoxSante.Hide()
        GroupBoxEducation.Hide()
        GroupBoxAgricultureEtElevage.Hide()
        GroupBoxAgricultureEtElevage.Hide()
        GroupBoxTransport.Hide()
        GroupBoxEquipementsMarchands.Hide()
        GroupBoxEnergie.Hide()
        GroupBoxSportsEtLoisirs.Hide()

        'GroupBoxFiltresGeographiques.Hide()
        'GroupBoxFiltresSpecifiquesToutInfras.Hide()

        'GenererRapportExcelInitial(sFile:="D:\Source\Clone de EDUCATION - labels - 2017-12-19-07-26.xlsx")
        'Dim sFile As String = "C:\Program RIC\Adolphe Expert\RIC\Source\Clone de EDUCATION - labels - 2017-12-22-08-36.xlsx"

        If Not Directory.Exists(sDirectory) Then
            Directory.CreateDirectory(sDirectory)
        End If
        If Not Directory.Exists(mDirectory) Then
            Directory.CreateDirectory(mDirectory)
        End If
        If Not Directory.Exists(rfDirectory) Then
            Directory.CreateDirectory(rfDirectory)
        End If
        If Not Directory.Exists(riDirectory) Then
            Directory.CreateDirectory(riDirectory)
        End If
        If Not File.Exists(sFileEducation) Then
            MsgBox("Veuillez charger les fichiers sources !", MsgBoxStyle.Exclamation, "Fichier source introuvable...")
            Exit Sub
        End If
        If Not File.Exists(mFileEducation) Then
            MsgBox("Veuillez charger les fichiers modèles !", MsgBoxStyle.Exclamation, "Fichier modèle introuvable...")
            Exit Sub
            'Else MsgBox("Voullez-vous chargez les rapports initiaux maintenant ?", MsgBoxStyle.YesNo, "Chargement rapport initaial...")
            '    'If MsgBoxResult.Yes = True Then
            '    GenererRapportExcelInitialEducation(sFileEducation)
            '    'End If
        End If

    End Sub

    Private Sub ButtonSelectionnerToutInfra_Click(sender As Object, e As EventArgs) Handles ButtonSelectionnerToutInfra.Click
        CheckBoxAdministration.Checked = True
        CheckBoxAgriculture.Checked = True
        CheckBoxEauAssainissement.Checked = True
        CheckBoxEducation.Checked = True
        CheckBoxElevage.Checked = True
        CheckBoxEnergie.Checked = True
        CheckBoxEquipementMarchand.Checked = True
        CheckBoxSante.Checked = True
        CheckBoxSportsLoisirs.Checked = True
        CheckBoxTransport.Checked = True

        'For Each Truc In GroupBox1.Controls
        '    Truc.Checked = True
        'Next Truc


    End Sub

    Private Sub ButtonReinitialiserInfra_Click(sender As Object, e As EventArgs) Handles ButtonReinitialiserInfra.Click
        CheckBoxAdministration.Checked = False
        CheckBoxAgriculture.Checked = False
        CheckBoxEauAssainissement.Checked = False
        CheckBoxEducation.Checked = False
        CheckBoxElevage.Checked = False
        CheckBoxEnergie.Checked = False
        CheckBoxEquipementMarchand.Checked = False
        CheckBoxSante.Checked = False
        CheckBoxSportsLoisirs.Checked = False
        CheckBoxTransport.Checked = False

    End Sub

    'Private Sub TreeViewFiltresGeographiques_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TreeViewFiltresGeographiques.AfterSelect
    '    If TreeViewFiltresGeographiques.SelectedNode.Checked = False Then
    '        TreeViewFiltresGeographiques.SelectedNode.Checked = True
    '    ElseIf TreeViewFiltresGeographiques.SelectedNode.Checked = True Then
    '        TreeViewFiltresGeographiques.SelectedNode.Checked = False
    '    End If
    'End Sub


    '***************************************

    ' Updates all child tree nodes recursively.
    Private Sub CheckAllChildNodes(treeNode As TreeNode, nodeChecked As Boolean)
        Dim node As TreeNode
        For Each node In treeNode.Nodes
            node.Checked = nodeChecked
            If node.Nodes.Count > 0 Then
                ' If the current node has child nodes, call the CheckAllChildsNodes method recursively.
                Me.CheckAllChildNodes(node, nodeChecked)
            End If
        Next node
    End Sub

    ' NOTE   This code can be added to the BeforeCheck event handler instead of the AfterCheck event.
    ' After a tree node's Checked property is changed, all its child nodes are updated to the same value.
    Private Sub node_AfterCheck(sender As Object, e As TreeViewEventArgs) Handles TreeViewFiltresGeographiques.AfterCheck
        ' The code only executes if the user caused the checked state to change.
        If e.Action <> TreeViewAction.Unknown Then
            If e.Node.Nodes.Count > 0 Then
                ' Calls the CheckAllChildNodes method, passing in the current 
                ' Checked value of the TreeNode whose checked state changed. 
                Me.CheckAllChildNodes(e.Node, e.Node.Checked)
            End If
        End If
    End Sub

    '///////////////////////////////////////

    Private Sub ComboBoxInfraAFiltrer_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxInfraAFiltrer.SelectedIndexChanged

        If ComboBoxInfraAFiltrer.Text = "Eau et Assainissement" Then
            Me.Controls.Add(GroupBoxEauEtAssainissement)
            GroupBoxEauEtAssainissement.Show()
            GroupBoxEauEtAssainissement.BringToFront()
            GroupBoxEauEtAssainissement.Visible = True
            Affichage = "Eau et Assainissement"
            '------------------------------------
            GroupBoxSportsEtLoisirs.Hide()
            GroupBoxAdministration.Hide()
            GroupBoxSante.Hide()
            GroupBoxEducation.Hide()
            GroupBoxAgricultureEtElevage.Hide()
            GroupBoxAgricultureEtElevage.Hide()
            GroupBoxTransport.Hide()
            GroupBoxEquipementsMarchands.Hide()
            GroupBoxEnergie.Hide()
            '------------------------------------
        ElseIf ComboBoxInfraAFiltrer.Text = "Sports et loisirs" Then
            Me.Controls.Add(GroupBoxSportsEtLoisirs)
            GroupBoxSportsEtLoisirs.Show()
            GroupBoxSportsEtLoisirs.BringToFront()
            GroupBoxSportsEtLoisirs.Visible = True
            Affichage = "Sports et loisirs"
            '------------------------------------
            GroupBoxEauEtAssainissement.Hide()
            GroupBoxAdministration.Hide()
            GroupBoxSante.Hide()
            GroupBoxEducation.Hide()
            GroupBoxAgricultureEtElevage.Hide()
            GroupBoxAgricultureEtElevage.Hide()
            GroupBoxTransport.Hide()
            GroupBoxEquipementsMarchands.Hide()
            GroupBoxEnergie.Hide()
            '------------------------------------
        ElseIf ComboBoxInfraAFiltrer.Text = "Administration" Then
            Me.Controls.Add(GroupBoxAdministration)
            GroupBoxAdministration.Show()
            GroupBoxAdministration.BringToFront()
            Affichage = "Administration"
            '------------------------------------
            GroupBoxEauEtAssainissement.Hide()
            GroupBoxSante.Hide()
            GroupBoxEducation.Hide()
            GroupBoxAgricultureEtElevage.Hide()
            GroupBoxAgricultureEtElevage.Hide()
            GroupBoxTransport.Hide()
            GroupBoxEquipementsMarchands.Hide()
            GroupBoxEnergie.Hide()
            GroupBoxSportsEtLoisirs.Hide()

            '------------------------------------
        ElseIf ComboBoxInfraAFiltrer.Text = "Santé" Then
            Me.Controls.Add(GroupBoxSante)
            GroupBoxSante.Show()
            GroupBoxSante.BringToFront()
            Affichage = "Santé"
            '------------------------------------
            GroupBoxAdministration.Hide()
            GroupBoxEauEtAssainissement.Hide()
            'GroupBoxSante.Hide()
            GroupBoxEducation.Hide()
            GroupBoxAgricultureEtElevage.Hide()
            GroupBoxAgricultureEtElevage.Hide()
            GroupBoxTransport.Hide()
            GroupBoxEquipementsMarchands.Hide()
            GroupBoxEnergie.Hide()
            GroupBoxSportsEtLoisirs.Hide()
            '------------------------------------
        ElseIf ComboBoxInfraAFiltrer.Text = "Education" Then
            Me.Controls.Add(GroupBoxEducation)
            GroupBoxEducation.Show()
            GroupBoxEducation.BringToFront()
            Affichage = "Education"
            '------------------------------------
            GroupBoxAdministration.Hide()
            GroupBoxEauEtAssainissement.Hide()
            GroupBoxSante.Hide()
            'GroupBoxEducation.Hide()
            GroupBoxAgricultureEtElevage.Hide()
            GroupBoxAgricultureEtElevage.Hide()
            GroupBoxTransport.Hide()
            GroupBoxEquipementsMarchands.Hide()
            GroupBoxEnergie.Hide()
            GroupBoxSportsEtLoisirs.Hide()
            '------------------------------------
        ElseIf ComboBoxInfraAFiltrer.Text = "Agriculture" Then
            Me.Controls.Add(GroupBoxAgricultureEtElevage)
            GroupBoxAgricultureEtElevage.Show()
            GroupBoxAgricultureEtElevage.BringToFront()
            Affichage = "Agriculture"
            '------------------------------------
            GroupBoxAdministration.Hide()
            GroupBoxEauEtAssainissement.Hide()
            GroupBoxSante.Hide()
            GroupBoxEducation.Hide()
            'GroupBoxAgricultureEtElevage.Hide()
            'GroupBoxAgricultureEtElevage.Hide()
            GroupBoxTransport.Hide()
            GroupBoxEquipementsMarchands.Hide()
            GroupBoxEnergie.Hide()
            GroupBoxSportsEtLoisirs.Hide()
            '------------------------------------

        ElseIf ComboBoxInfraAFiltrer.Text = "Elevage" Then
            Me.Controls.Add(GroupBoxAgricultureEtElevage)
            GroupBoxAgricultureEtElevage.Show()
            GroupBoxAgricultureEtElevage.BringToFront()
            Affichage = "Elevage"
            '------------------------------------
            GroupBoxAdministration.Hide()
            GroupBoxEauEtAssainissement.Hide()
            GroupBoxSante.Hide()
            GroupBoxEducation.Hide()
            'GroupBoxAgricultureEtElevage.Hide()
            'GroupBoxAgricultureEtElevage.Hide()
            GroupBoxTransport.Hide()
            GroupBoxEquipementsMarchands.Hide()
            GroupBoxEnergie.Hide()
            GroupBoxSportsEtLoisirs.Hide()
            '------------------------------------

        ElseIf ComboBoxInfraAFiltrer.Text = "Transport" Then
            Me.Controls.Add(GroupBoxTransport)
            GroupBoxTransport.Show()
            GroupBoxTransport.BringToFront()
            Affichage = "Transport"
            '------------------------------------
            GroupBoxAdministration.Hide()
            GroupBoxEauEtAssainissement.Hide()
            GroupBoxSante.Hide()
            GroupBoxEducation.Hide()
            GroupBoxAgricultureEtElevage.Hide()
            GroupBoxAgricultureEtElevage.Hide()
            'GroupBoxTransport.Hide()
            GroupBoxEquipementsMarchands.Hide()
            GroupBoxEnergie.Hide()
            GroupBoxSportsEtLoisirs.Hide()
            '------------------------------------

        ElseIf ComboBoxInfraAFiltrer.Text = "Equipements marchands" Then
            Me.Controls.Add(GroupBoxEquipementsMarchands)
            GroupBoxEquipementsMarchands.Show()
            GroupBoxEquipementsMarchands.BringToFront()
            Affichage = "Equipements marchands"

            '------------------------------------
            GroupBoxAdministration.Hide()
            GroupBoxEauEtAssainissement.Hide()
            GroupBoxSante.Hide()
            GroupBoxEducation.Hide()
            GroupBoxAgricultureEtElevage.Hide()
            GroupBoxAgricultureEtElevage.Hide()
            GroupBoxTransport.Hide()
            'GroupBoxEquipementsMarchands.Hide()
            GroupBoxEnergie.Hide()
            GroupBoxSportsEtLoisirs.Hide()
            '------------------------------------

        ElseIf ComboBoxInfraAFiltrer.Text = "Energie" Then
            Me.Controls.Add(GroupBoxEnergie)
            GroupBoxEnergie.Show()
            GroupBoxEnergie.BringToFront()
            Affichage = "Energie"

            '------------------------------------
            GroupBoxAdministration.Hide()
            GroupBoxEauEtAssainissement.Hide()
            GroupBoxSante.Hide()
            GroupBoxEducation.Hide()
            GroupBoxAgricultureEtElevage.Hide()
            GroupBoxAgricultureEtElevage.Hide()
            GroupBoxTransport.Hide()
            GroupBoxEquipementsMarchands.Hide()
            'GroupBoxEnergie.Hide()
            GroupBoxSportsEtLoisirs.Hide()
            '------------------------------------
        End If
    End Sub
    '////////////////////////////////////////////

    Private Sub ButtonSelectAllFiltres_Click(sender As Object, e As EventArgs) Handles ButtonSelectAllFiltres.Click
        If Affichage = "Education" Then
            CheckBoxEcolePrive.Checked = True
            CheckBoxEcolePublique.Checked = True
            CheckBoxBonEtatEducation.Checked = True
            CheckBoxEtatUsageEducation.Checked = True
            CheckBoxInutilisableEducation.Checked = True
            CheckBoxPrescolaire.Checked = True
            CheckBoxPrimaire.Checked = True
            CheckBoxCollege.Checked = True
            CheckBoxLycee.Checked = True
            NombreSup = NumericUpDown1.Value


        End If
    End Sub

    Private Sub ButtonCancelFiltres_Click(sender As Object, e As EventArgs) Handles ButtonCancelFiltres.Click
        If Affichage = "Education" Then
            CheckBoxEcolePrive.Checked = False
            CheckBoxEcolePublique.Checked = False
            CheckBoxBonEtatEducation.Checked = False
            CheckBoxEtatUsageEducation.Checked = False
            CheckBoxInutilisableEducation.Checked = False
            CheckBoxPrescolaire.Checked = False
            CheckBoxPrimaire.Checked = False
            CheckBoxCollege.Checked = False
            CheckBoxLycee.Checked = False
            NombreSup = 0
            NumericUpDown1.Value = 0

        End If
    End Sub

    Private Sub QuitterToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles QuitterToolStripMenuItem.Click
        Me.Close()
    End Sub

    'Private Sub ButtonVoirLesFiltres_Click(sender As Object, e As EventArgs) Handles ButtonVoirLesFiltres.Click
    '    'GroupBoxFiltresGeographiques.Show()
    '    'GroupBoxFiltresSpecifiquesToutInfras.Show()
    '    ComboBoxInfraAFiltrer.Items.Clear()
    '    ComboBoxInfraAFiltrer.Text = ""
    '    If CheckBoxAdministration.Checked = True Then
    '        ComboBoxInfraAFiltrer.Items.Add("Administration")
    '    Else
    '        ComboBoxInfraAFiltrer.Items.Remove("Administration")
    '    End If
    '    If CheckBoxEducation.Checked = True Then
    '        ComboBoxInfraAFiltrer.Items.Add("Education")
    '    Else
    '        ComboBoxInfraAFiltrer.Items.Remove("Education")
    '    End If
    '    If CheckBoxSportsLoisirs.Checked = True Then
    '        ComboBoxInfraAFiltrer.Items.Add("Sports et loisirs")
    '    Else
    '        ComboBoxInfraAFiltrer.Items.Remove("Sports et loisirs")
    '    End If
    '    If CheckBoxEauAssainissement.Checked = True Then
    '        ComboBoxInfraAFiltrer.Items.Add("Eau et Assainissement")
    '    Else
    '        ComboBoxInfraAFiltrer.Items.Remove("Eau et Assainissement")
    '    End If
    '    If CheckBoxSante.Checked = True Then
    '        ComboBoxInfraAFiltrer.Items.Add("Santé")
    '    Else
    '        ComboBoxInfraAFiltrer.Items.Remove("Santé")
    '    End If
    '    If CheckBoxAgriculture.Checked = True Then
    '        ComboBoxInfraAFiltrer.Items.Add("Agriculture")
    '    Else
    '        ComboBoxInfraAFiltrer.Items.Remove("Agriculture")
    '    End If
    '    If CheckBoxElevage.Checked = True Then
    '        ComboBoxInfraAFiltrer.Items.Add("Elevage")
    '    Else
    '        ComboBoxInfraAFiltrer.Items.Remove("Elevage")
    '    End If
    '    If CheckBoxTransport.Checked = True Then
    '        ComboBoxInfraAFiltrer.Items.Add("Transport")
    '    Else
    '        ComboBoxInfraAFiltrer.Items.Remove("Transport")
    '    End If
    '    If CheckBoxEquipementMarchand.Checked = True Then
    '        ComboBoxInfraAFiltrer.Items.Add("Equipements marchands")
    '    Else
    '        ComboBoxInfraAFiltrer.Items.Remove("Equipements marchands")
    '    End If
    '    If CheckBoxEnergie.Checked = True Then
    '        ComboBoxInfraAFiltrer.Items.Add("Energie")
    '    Else
    '        ComboBoxInfraAFiltrer.Items.Remove("Energie")
    '    End If
    'End Sub

    Private Sub CheckBoxAdministration_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxAdministration.CheckedChanged
        'ComboBoxInfraAFiltrer.Items.Clear()
        'ComboBoxInfraAFiltrer.Text = ""
        If CheckBoxAdministration.Checked = True Then
            ComboBoxInfraAFiltrer.Items.Add("Administration")
        Else
            ComboBoxInfraAFiltrer.Items.Remove("Administration")
        End If
    End Sub

    Private Sub CheckBoxAgriculture_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxAgriculture.CheckedChanged
        If CheckBoxAgriculture.Checked = True Then
            ComboBoxInfraAFiltrer.Items.Add("Agriculture")
        Else
            ComboBoxInfraAFiltrer.Items.Remove("Agriculture")
        End If
    End Sub

    Private Sub CheckBoxEauAssainissement_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxEauAssainissement.CheckedChanged
        If CheckBoxEauAssainissement.Checked = True Then
            ComboBoxInfraAFiltrer.Items.Add("Eau et Assainissement")
        Else
            ComboBoxInfraAFiltrer.Items.Remove("Eau et Assainissement")
        End If
    End Sub

    Private Sub CheckBoxElevage_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxElevage.CheckedChanged
        If CheckBoxElevage.Checked = True Then
            ComboBoxInfraAFiltrer.Items.Add("Elevage")
        Else
            ComboBoxInfraAFiltrer.Items.Remove("Elevage")
        End If
    End Sub

    Private Sub CheckBoxEducation_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxEducation.CheckedChanged
        If CheckBoxEducation.Checked = True Then
            ComboBoxInfraAFiltrer.Items.Add("Education")
            ComboBoxInfraAFiltrer.SelectedItem = "Education"
        Else
            ComboBoxInfraAFiltrer.Items.Remove("Education")
            ComboBoxInfraAFiltrer.SelectedItem = ""
            GroupBoxEducation.Visible = False
        End If
    End Sub

    Private Sub CheckBoxEquipementMarchand_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxEquipementMarchand.CheckedChanged
        If CheckBoxEquipementMarchand.Checked = True Then
            ComboBoxInfraAFiltrer.Items.Add("Equipements marchands")
        Else
            ComboBoxInfraAFiltrer.Items.Remove("Equipements marchands")
        End If
    End Sub

    Private Sub CheckBoxEnergie_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxEnergie.CheckedChanged
        If CheckBoxEnergie.Checked = True Then
            ComboBoxInfraAFiltrer.Items.Add("Energie")
        Else
            ComboBoxInfraAFiltrer.Items.Remove("Energie")
        End If
    End Sub

    Private Sub CheckBoxSante_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxSante.CheckedChanged
        If CheckBoxSante.Checked = True Then
            ComboBoxInfraAFiltrer.Items.Add("Santé")
        Else
            ComboBoxInfraAFiltrer.Items.Remove("Santé")
        End If

    End Sub

    Private Sub CheckBoxSportsLoisirs_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxSportsLoisirs.CheckedChanged
        If CheckBoxSportsLoisirs.Checked = True Then
            ComboBoxInfraAFiltrer.Items.Add("Sports et loisirs")
        Else
            ComboBoxInfraAFiltrer.Items.Remove("Sports et loisirs")
        End If
    End Sub

    Private Sub CheckBoxTransport_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxTransport.CheckedChanged
        If CheckBoxTransport.Checked = True Then
            ComboBoxInfraAFiltrer.Items.Add("Transport")
        Else
            ComboBoxInfraAFiltrer.Items.Remove("Transport")

            'ComboBoxInfraAFiltrer.Items.Add("Transport")
            'ComboBoxInfraAFiltrer.Enabled = False
            'ComboBoxInfraAFiltrer.SelectedItem = False

        End If
    End Sub

    Private Sub TreeViewFiltresGeographiques_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TreeViewFiltresGeographiques.AfterSelect

    End Sub

    Private Sub EducationToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EducationToolStripMenuItem.Click
        ' OpenFileDialogEducation.ShowDialog()
        With OpenFileDialogEducation
            .Title = "Fichier excel source pour le RIC - Infrastructure ''EDUCATION'' "           ' DIALOG BOX TITLE.
            .FileName = ""
            .Filter = "Fichier Excel RIC|*.xlsx;*.xls"     ' FILTER ONLY EXCEL FILES IN FILE TYPE.

            If .ShowDialog() = DialogResult.OK Then
                sFileName = .FileName

                'If Trim(sFileName) <> "" Then
                '    EditEmpDetails(sFileName)       ' PROCEDURE TO EDIT EMPLOYEE DETAILS.
                'End If
                'LabelSourceEducation.Text = OpenFileDialogEducation.FileName

                'path = OpenFileDialogEducation.FileName
                sFileI = OpenFileDialogEducation.FileName
                If Not File.Exists("C:\Program RIC\Adolphe Expert\RIC\Source\" & sFileI & ".xlsx") Then
                    MsgBox("Voullez-vous chargez les fichiers sources maintenant ?", MsgBoxStyle.OkCancel, "Fichier source introuvable...")
                    'If MsgBoxResult.Ok = True Then
                    File.Copy(sFileI, "C:\Program RIC\Adolphe Expert\RIC\Source\" & sFileI & ".xlsx")
                    If Not File.Exists("C:\Program RIC\Adolphe Expert\RIC\Source\" & sFileI & ".xlsx") Then
                        MsgBox("Veuillez réessayer plus-tard...", MsgBoxStyle.Exclamation, "Fichier source non copié !")
                        Exit Sub
                    End If
                    'Else
                    '        Exit Sub
                    'End If
                    ' Exit Sub
                Else
                    MsgBox("Voullez-vous changez les fichiers sources ?", MsgBoxStyle.OkCancel, "Fichier source existe dejà...")
                    Exit Sub
                End If

            End If

        End With
    End Sub

    'Private Sub LoadData()
    '    con = New OleDbConnection(conString)
    '    Dim query As String = "SELECT * FROM [EDUCATION$] "
    '    adapter = New OleDbDataAdapter(query, con)

    '    Dim ds As DataSet = New DataSet()
    '    Dim dt As DataTable = New DataTable

    '    adapter.Fill(dt)

    '    DataGridView1.DataSource = ds.Tables(0)
    '    DataGridView1.DataMember = "[EDUCATION$]"

    '    con.Close()

    '    '------------------------------------------

    '    ''déclaration du dataset
    '    'Dim dat As DataSet
    '    'dat = New DataSet
    '    ''déclaration et utilisation d'un OLeDBConnection
    '    'Using Conn As OleDbConnection = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='D:\EDUCATION - labels - 2017-11-22-14-22.xlsx';Extended Properties=""Excel 8.0;HDR=Yes;""")
    '    '    ' Conn.Open()
    '    '    'déclaration du DataAdapter
    '    '    'notre requête sélectionne toute les cellule de la Feuil1
    '    '    Using Adap As OleDbDataAdapter = New OleDbDataAdapter("select * from [EDUCATION$]", Conn)
    '    '        'Chargement du Dataset
    '    '        Adap.Fill(dat)
    '    '        'On Binde les données sur le DGV
    '    '        DataGridView1.DataSource = dat.Tables(0)
    '    '    End Using
    '    '    'le end using libère les ressources
    '    'End Using
    '    '------------------------------------------

    'End Sub

    Private Sub ButtonAppliquerFiltres_Click(sender As Object, e As EventArgs) Handles ButtonAppliquerFiltres.Click

        '1st method start /////////////////////////////////////////////////
        Try
            Dim MyConnection As System.Data.OleDb.OleDbConnection
            Dim DataSet As System.Data.DataSet
            Dim MyCommand As System.Data.OleDb.OleDbDataAdapter

            'Dim path As String = "D:\EDUCATION - labels - 2017-11-22-14-22.xlsx"

            'Dim path As String = LabelSourceEducation.Text
            'LabelSourceEducation.Text definit la source des données

            'Dim sourceDesDonnees As String = path

            'If LabelSourceEducation.Text = "" Then

            'End If


            MyConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + sFileEducation + ";Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1;""")
            'MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [EDUCATION$] Where Eau_potable_accesssible_sur_le = 'Oui' ", MyConnection)
            'MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [EDUCATION$]", MyConnection)
            'MyCommand = New System.Data.OleDb.OleDbDataAdapter("select [DISTRICT], Nom_de_l_tablissement_scolaire, Type_d_tablissement, Niveaux_enseign_s from [EDUCATION$]", MyConnection)
            'MyCommand = New System.Data.OleDb.OleDbDataAdapter("select [_index], [DISTRICT], [Nom_de_l_tablissement_scolaire], [Type_d_tablissement], [Niveaux_enseign_s] from [EDUCATION$]", MyConnection)
            'MyCommand = New System.Data.OleDb.OleDbDataAdapter("select [Préciser le genre d'Infrastructure Communale], [Préciser le nom UNIQUE de cette infrastructure], [DISTRICT dans lequel se trouve l'infrastructure], [Nom de l'établissement scolaire], [COMMUNE (District Diégo)] from [Clone de EDUCATION$]", MyConnection)
            MyCommand = New System.Data.OleDb.OleDbDataAdapter("select [Préciser le genre d'Infrastructure Communale], [Préciser le nom UNIQUE de cette infrastructure], [COMMUNE (District Diégo)], [DISTRICT dans lequel se trouve l'infrastructure], [Nom de l'établissement scolaire] from [Clone de EDUCATION$]", MyConnection)

            DataSet = New System.Data.DataSet

            MyCommand.Fill(DataSet)

            DataGridView1.DataSource = DataSet.Tables(0)

            'DataGridView1.ColumnCount = 7

            Dim chk As New DataGridViewCheckBoxColumn()
            'Dim gridButtom As New DataGridViewLinkColumn

            DataGridView1.Columns.Add(chk)

            chk.HeaderText = "Cocher pour voir détails"

            chk.Name = "chk"

            'Dim gridButtom As New DataGridViewButtonColumn

            'DataGridView1.Columns.Add(gridButtom)

            'gridButtom.HeaderText = "gridButtom"

            'gridButtom.Name = "gridButtom"

            'DataGridView1.Columns(0).HeaderText = "Entête-1"
            'DataGridView1.Columns(1).HeaderText = "Entête-2"
            'DataGridView1.Columns(2).HeaderText = "Entête-3"
            'DataGridView1.Columns(3).HeaderText = "Entête-4"
            'DataGridView1.Columns(4).HeaderText = "Entête-5"


            'DataGridView1.Rows(2).Cells(3).Value = True
            MyConnection.Close()
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try
        '1st method end ////////////////////////////////////////////////////

        'LoadData()

        '2nd method start ////////////////////////////////////////////////////

        'Dim con As New OleDbConnection
        'Dim cm As New OleDbCommand
        'con = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Documents and Settings\crysol\Desktop\TEST\Book1.xls;Extended Properties=""Excel 12.0 Xml;HDR=YES""")
        'con.Open()
        'With cm
        '    .Connection = con
        '    .CommandText = "update [up$] set [name]=?, [QC_status]=?, [reason]=?, [date]=? WHERE [article_no]=?"
        '    cm = New OleDbCommand(.CommandText, con)
        '    cm.Parameters.AddWithValue("?", TextBox2.Text)
        '    cm.Parameters.AddWithValue("?", ComboBox1.SelectedItem)
        '    cm.Parameters.AddWithValue("?", TextBox3.Text)
        '    cm.Parameters.AddWithValue("?", DateTimePicker1.Text)
        '    cm.Parameters.AddWithValue("?", TextBox1.Text)
        '    cm.ExecuteNonQuery()
        '    MsgBox("UPDATE SUCCESSFUL")
        '    con.Close()
        'End With

        '2nd method end  /////////////////////////////////////////////////////

    End Sub

    'Private Sub ButtonVoirRapportDetaille_Click(sender As Object, e As EventArgs) Handles ButtonVoirRapportDetaille.Click

    '///////////////////'1st method start

    'Try
    '    Dim oXL As Excel.Application
    '    Dim oWB As Excel.Workbook
    '    Dim oSheet As Excel.Worksheet
    '    Dim oRng As Excel.Range
    '    'On Error GoTo Err_Handler
    '    ' Start Excel and get Application object.
    '    oXL = New Excel.Application

    '    'Get a new workbook.
    '    Dim path As String = ViewState("filepath")
    '    oWB = oXL.Workbooks.Open(path)
    '    oSheet = CType(oXL.Worksheets.Add(Missing.Value, Missing.Value, Missing.Value, Missing.Value), Excel.Worksheet)
    '    'oSheet.Name = "Reject_History"
    '    Dim totalSheets As Integer = oXL.ActiveWorkbook.Sheets.Count
    '    CType(oXL.ActiveSheet, Excel.Worksheet).Move(After:=oXL.Worksheets(totalSheets))
    '    CType(oXL.ActiveWorkbook.Sheets(totalSheets), Excel.Worksheet).Activate()

    '    'Write Dataset to Excel Sheet
    '    Dim col As Integer = 0

    '    For Each dr As DataColumn In DirectCast(ViewState("DisplayNonExisting"), DataTable).Columns

    '        col += 1
    '        'Determine cell to write
    '        oSheet.Cells(10, col).Value = dr.ColumnName

    '    Next

    '    Dim irow As Integer = 10
    '    For Each dr As DataRow In DirectCast(ViewState("DisplayNonExisting"), DataTable).Rows
    '        irow += 1
    '        Dim icol As Integer = 0
    '        For Each c As String In dr.ItemArray
    '            icol += 1
    '            'Determine cell to write
    '            oSheet.Cells(irow, icol).Value = c
    '        Next
    '    Next

    '    ' Make sure Excel is visible and give the user control
    '    ' of Microsoft Excel's lifetime.
    '    ' oXL.Visible = True
    '    ' oXL.UserControl = True

    '    'oWB.SaveAs(path, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, False, False, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing)
    '    'oWB.Close()
    '    oWB.Save()
    '    oWB.Close(Type.Missing, Type.Missing, Type.Missing)
    '    ' Make sure you release object references.

    '    oRng = Nothing
    '    oSheet = Nothing
    '    oWB = Nothing
    '    oXL = Nothing

    'Catch ex As Exception
    'MsgBox(ex.Message.ToString)
    'End Try
    '///////////////////'1st method end
    'End Sub

    Private Sub ButtonVoirRapportDetaille_Click(sender As Object, e As EventArgs) Handles ButtonVoirRapportDetaille.Click

        '*************************** 2nd method
        Try
            '~~> Define your Excel Objects
            Dim xlApp As New Excel.Application
            Dim xlWorkBook As Excel.Workbook
            'Dim xlWorkSheet As Excel.Worksheet

            '~~> Opens Source Workbook. Change path and filename as applicable
            'xlWorkBook = xlApp.Workbooks.Open("D:\Source\Clone de EDUCATION - labels - 2017-12-19-07-26.xlsx")
            If Not File.Exists("C:\Program RIC\Adolphe Expert\RIC\Rapport Initial\" & nRapportEducation & "- District de " & nDistrictEducation & ".xlsx") Then
                MsgBox("Fichier" & nRapportEducation & "- District de " & nDistrictEducation & ".xlsx est introuvable, Veuillez charger les fichiers rapports initiaux !", MsgBoxStyle.Exclamation, "Fichier introuvable...")
                Exit Sub
            End If
            xlWorkBook = xlApp.Workbooks.Open("C:\Program RIC\Adolphe Expert\RIC\Rapport Initial\" & nRapportEducation & "- District de " & nDistrictEducation & ".xlsx")

            'xlWorkBook = xlApp.Workbooks.Open("D:\Rapport Initial\CRÈCHE ANKETRABE I.xlsx")
            xlApp.Visible = True
            'xlWorkBookTest = xlAppTest.Workbooks.Open("D:\Rapport Initial\" & nRapport & ".xlsx")

            'System.Diagnostics.Process.Start("D:\Rapport Initial\" & nRapport & ".xlsx")

        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try

        '*************************** 2nd method
    End Sub
    ''' <summary>
    ''' Generer le fichier excel de rapport initial pour l'infrastructure Education
    ''' </summary>
    ''' <param name="sFileEducation"></param>
    Private Sub GenererRapportExcelInitialEducation(ByVal sFileEducation As String)

        Dim nLigne As Integer = 2
        'Dim fLigne As Integer = nLigne + 1

        For nLigne = 2 To 100

            '~~> Define your Excel Objects
            Dim xlApp As New Excel.Application
            Dim xlWorkBook, xlWorkBook2 As Excel.Workbook
            Dim xlWorkSheet, xlWsheet2 As Excel.Worksheet
            'Dim xlSourceRange, xlDestRange As Excel.Range
            'Dim misValue As Object = System.Reflection.Missing.Value

            'Try

            '~~> Opens Source Workbook. Change path and filename as applicable
            'xlWorkBook = xlApp.Workbooks.Open("D:\Source\Clone de EDUCATION - labels - 2017-12-19-07-26.xlsx")
            xlWorkBook = xlApp.Workbooks.Open(sFileEducation)

            '~~> Opens Destination Workbook. Change path and filename as applicable
            'xlWorkBook2 = xlApp.Workbooks.Open("D:\Model\Education v1.xlsx")
            xlWorkBook2 = xlApp.Workbooks.Open(mFileEducation)

            '~~> Display Excel
            xlApp.Visible = False

            'xlWorkBook = xlApp.Workbooks.Add(misValue)
            'xlWorkBook = xlApp.Workbook2.Add(misValue)

            '~~> Set the source worksheet
            xlWorkSheet = xlWorkBook.Sheets("Clone de EDUCATION")
            '~~> Set the destination worksheet
            xlWsheet2 = xlWorkBook2.Sheets("Feuil1")

            'For nLigne = CInt(xlWorkSheet.Range("UI2").Text) To CInt(xlWorkSheet.Range("UI200").Text)
            'For nLigne = xlWorkSheet.Range("UI2").Value To xlWorkSheet.Range("UI200").Value

            If xlWorkSheet.Range("A" & nLigne).Value <> "" Then

                'insert pic
                'xlWsheet2.Range("E6").Shapes.AddPicture("C:\butterfly.jpg",
                xlWsheet2.Shapes.AddPicture("C:\butterfly.jpg",
             Microsoft.Office.Core.MsoTriState.msoFalse,
             Microsoft.Office.Core.MsoTriState.msoCTrue, 280, 80, 200, 150)

                xlWsheet2.Shapes.AddPicture("C:\Baobab.jpg",
             Microsoft.Office.Core.MsoTriState.msoFalse,
             Microsoft.Office.Core.MsoTriState.msoCTrue, 485, 80, 200, 150)

                'nRapport = xlWorkSheet.Range("E2").Text
                nRapportEducation = xlWorkSheet.Range("E" & nLigne).Text
                nDistrictEducation = xlWorkSheet.Range("F" & nLigne).Text

                MsgBox("Maintenant la ligne : " & nLigne, MsgBoxStyle.Exclamation, "Vérification des lignes d'enquentes...")

                xlWsheet2.Range("E2").Value = xlWorkSheet.Range("E" & nLigne).Text   'Préciser le nom unique de cette infrastructure
                xlWsheet2.Range("E1").Value = xlWorkSheet.Range("D" & nLigne).Text   'Préciser le genre d'infrastructure communale

                '***************************************************************************************

                Dim SautLigne As Integer = 0

                'Dim sautLigne As Integer
                ''~~> Set the source range
                'For sautLigne = 6 To 18
                xlWsheet2.Range("B6").Value = xlWorkSheet.Range("F" & nLigne).Text   'Disctrict dans lequel se trouve l'infrastructure
                'Next sautLigne
                ''xlSourceRange = xlWorkSheet.Range("F" & index1 + 1)
                ''xlSourceRange = xlWorkSheet.Range("F" & index1.ToString + 1)
                ''~~> Set the destination range
                'xlDestRange = xlWsheet2.Range("B6")
                ''~~> Copy and paste the range
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("B7").Value = xlWorkSheet.Range("G" & nLigne).Text  'Commune (District Diégo)
                'xlDestRange = xlWsheet2.Range("B7")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("B8").Value = xlWorkSheet.Range("H" & nLigne).Text   'Commune (District Ambilobe)
                'xlDestRange = xlWsheet2.Range("B8")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("B9").Value = xlWorkSheet.Range("I" & nLigne).Text   'Commune (District Ambanja)
                'xlDestRange = xlWsheet2.Range("B9")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("B10").Value = xlWorkSheet.Range("J" & nLigne).Text    'Nom du fokontany (Commune Joffre ville)
                'xlDestRange = xlWsheet2.Range("B10")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("B11").Value = xlWorkSheet.Range("K" & nLigne).Text     'Nom du fokontany (Commune Sakaramy)
                'xlDestRange = xlWsheet2.Range("B11")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("B12").Value = xlWorkSheet.Range("L" & nLigne).Text    'Nom du fokontany (Commune Antsahampano)
                'xlDestRange = xlWsheet2.Range("B12")
                'xlSourceRange.Copy(xlDestRange)
                xlWsheet2.Range("B13").Value = xlWorkSheet.Range("M" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("M" & nLigne)    'Nom du fokontany (Commune Mantaly)
                'xlDestRange = xlWsheet2.Range("B13")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("B14").Value = xlWorkSheet.Range("N" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("N" & nLigne)    'Nom du fokontany (Commune Marivorahona)
                'xlDestRange = xlWsheet2.Range("B14")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("B15").Value = xlWorkSheet.Range("O" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("O" & nLigne)    'Nom du fokontany (Commune Ampondralava)
                'xlDestRange = xlWsheet2.Range("B15")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("B16").Value = xlWorkSheet.Range("P" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("P" & nLigne)    'Nom du fokontany (Commune Antsakoamanondro)
                'xlDestRange = xlWsheet2.Range("B16")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("B17").Value = xlWorkSheet.Range("Q" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("Q" & nLigne)    'Nom du fokontany (Commune Ambohimena)
                'xlDestRange = xlWsheet2.Range("B17")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("B18").Value = xlWorkSheet.Range("R" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("R" & nLigne)    'Nom du fokontany (Commune Antranokarany)
                'xlDestRange = xlWsheet2.Range("B18")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("F24").Value = xlWorkSheet.Range("S" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("S" & nLigne)    'Nom de l'établissement scolaire
                'xlDestRange = xlWsheet2.Range("F24")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("F25").Value = xlWorkSheet.Range("T" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("T" & nLigne) 'Localisation GPS de l'établissement
                'xlDestRange = xlWsheet2.Range("F25")
                'xlSourceRange.Copy(xlDestRange)

                'xlWsheet2.Range("D15").Value = xlWorkSheet.Range("U" & nLigne).Text
                ''xlSourceRange = xlWorkSheet.Range("U" & nLigne) '_Localisation GPS de l'établissement_latitude
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("V" & nLigne)  '_Localisation GPS de l'établissement_longitude
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("W" & nLigne)  '_Localisation GPS de l'établissement_altitude
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("X" & nLigne)  '_Localisation GPS de l'établissement_precision
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)
                ''xlWsheet2.Range("D15").Font.Size = 12

                xlWsheet2.Range("F26").Value = xlWorkSheet.Range("Y" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("Y" & nLigne)  'Type d'établissement
                'xlDestRange = xlWsheet2.Range("F26")
                'xlSourceRange.Copy(xlDestRange)
                ''xlWsheet2.Range("F26").Font.Size = 12

                xlWsheet2.Range("F27").Value = xlWorkSheet.Range("Z" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("Z" & nLigne)  'Niveaux enseignés
                'xlDestRange = xlWsheet2.Range("F27")
                'xlSourceRange.Copy(xlDestRange)
                ''xlWsheet2.Range("F27").Font.Size = 12

                'xlWsheet2.Range("B14").Value = xlWorkSheet.Range("N" & nLigne).Text
                ''xlSourceRange = xlWorkSheet.Range("AA" & nLigne)  'Niveaux enseignés/Préscolaire
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("AB" & nLigne)  'Niveaux enseignés/Primaire (11ème - 10 ème - 9ème - 8ème - 7ème)
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("AC" & nLigne)   'Niveaux enseignés/Collège (6ème - 5ème - 4ème - 3ème)
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("AD" & nLigne)   'Niveaux enseignés/Lycée (2nde - 1ère - Terminale)
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("F28").Value = xlWorkSheet.Range("AE" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("AE" & nLigne)   'Construction de l'établissement scolaire financée par
                'xlDestRange = xlWsheet2.Range("F28")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("F29").Value = xlWorkSheet.Range("AF" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("AF" & nLigne)  'Si autre, préciser
                'xlDestRange = xlWsheet2.Range("F29")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("F30").Value = xlWorkSheet.Range("AG" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("AG" & nLigne)   'Coût de construction de l'établissement scolaire (en Ariary)
                'xlDestRange = xlWsheet2.Range("F30")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("G31").Value = xlWorkSheet.Range("AH" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("AH" & nLigne)  'Type de participation de la commune / des usagers à la construction
                'xlDestRange = xlWsheet2.Range("G31")
                'xlSourceRange.Copy(xlDestRange)

                'xlWsheet2.Range("D15").Value = xlWorkSheet.Range("AI" & nLigne).Text
                ''xlSourceRange = xlWorkSheet.Range("AI" & nLigne)   'Type de participation de la commune / des usagers à la construction/Financière
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("AJ" & nLigne)   'Type de participation de la commune / des usagers à la construction/Autre
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("AK" & nLigne)   'Type de participation de la commune / des usagers à la construction/Main d'oeuvre
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("AL" & nLigne)  'Type de participation de la commune / des usagers à la construction/Aucune participation
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("AM" & nLigne)   'Type de participation de la commune / des usagers à la construction/Terrain
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("G32").Value = xlWorkSheet.Range("AN" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("AN" & nLigne)   'Si autre, préciser le type de participation
                'xlDestRange = xlWsheet2.Range("G32")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("H33").Value = xlWorkSheet.Range("AO" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("AO" & nLigne)   'Si participation financière, préciser le montant (en Ariary)
                'xlDestRange = xlWsheet2.Range("H33")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("G34").Value = xlWorkSheet.Range("AP" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("AP" & nLigne)   'Préciser la situation foncière du terrain sur lequel est située l'infrastructure
                'xlDestRange = xlWsheet2.Range("G34")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("G35").Value = xlWorkSheet.Range("AQ" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("AQ" & nLigne)
                'xlDestRange = xlWsheet2.Range("G35")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("F36").Value = xlWorkSheet.Range("AR" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("AR" & nLigne)   'Préciser qui est le propiétaire de l'infrastructure (Etat / Commune / Privé / ...)
                'xlDestRange = xlWsheet2.Range("F36")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("F37").Value = xlWorkSheet.Range("AS" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("AS" & nLigne)   'Année de construction de l'établissement scolaire (préciser si différentes dates de construction)
                'xlDestRange = xlWsheet2.Range("F37")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("F38").Value = xlWorkSheet.Range("AT" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("AT" & nLigne)   'Etat général de l'établissement scolaire
                'xlDestRange = xlWsheet2.Range("F38")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("I39").Value = xlWorkSheet.Range("AU" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("AU" & nLigne)  'Décrire rapidement les raisons qui rendent l’infrastructure non fonctionnelle
                'xlDestRange = xlWsheet2.Range("I39")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("I40").Value = xlWorkSheet.Range("AV" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("AV" & nLigne)   'Le dimensionnement de l’infrastructure est-il en adéquation avec le nombre d’usagers ?
                'xlDestRange = xlWsheet2.Range("I40")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("I41").Value = xlWorkSheet.Range("AW" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("AW" & nLigne)   'Le dimensionnement de l’infrastructure doit-il être prochainement revu en raison de la prévision d’une augmentation significative de la population dans un futur proche ?
                'xlDestRange = xlWsheet2.Range("I41")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("I45").Value = xlWorkSheet.Range("AX" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("AX" & nLigne)   'Des équipements sportifs sont-ils disponibles sur le site ?
                'xlDestRange = xlWsheet2.Range("I45")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("F46").Value = xlWorkSheet.Range("AY" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("AY" & nLigne)   'Si oui, préciser le type d'équipement sportif
                'xlDestRange = xlWsheet2.Range("F46")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("G47").Value = xlWorkSheet.Range("AZ" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("AZ" & nLigne)    'Description technique succinte de l'équipement sportif (dimension, matériaux,...)
                'xlDestRange = xlWsheet2.Range("G47")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("G48").Value = xlWorkSheet.Range("BA" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("BA" & nLigne)   'Préciser l'état général de l'équipement sportif
                'xlDestRange = xlWsheet2.Range("G48")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D52").Value = xlWorkSheet.Range("BB" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("BB" & nLigne)   'Eau accesssible sur le site de l'établissement scolaire
                'xlDestRange = xlWsheet2.Range("D52")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D53").Value = xlWorkSheet.Range("BC" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("BC" & nLigne)   'Infrastructure fournissant l'eau
                'xlDestRange = xlWsheet2.Range("D53")
                'xlSourceRange.Copy(xlDestRange)

                'xlWsheet2.Range("D15").Value = xlWorkSheet.Range("BD" & nLigne).Text
                ''xlSourceRange = xlWorkSheet.Range("BD" & nLigne)   'Infrastructure fournissant l'eau/Puit protégé (couvert afin de limiter les risques de contamination)
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("BE" & nLigne)   'Infrastructure fournissant l'eau/Autre
                ''xlDestRange = xlWsheet2.Range("D54")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("BF" & nLigne)   'Infrastructure fournissant l'eau/Puit ouvert
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("BG" & nLigne)   'Infrastructure fournissant l'eau/Connexion à JIRAMA
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("BH" & nLigne)   'Infrastructure fournissant l'eau/Borne fontaine
                ''xlDestRange = xlWsheet2.Range("D55")
                ''xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D54").Value = xlWorkSheet.Range("BI" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("BI" & nLigne)   'Si autre, préciser la nature de l'infrastructure fournissant l'eau
                'xlDestRange = xlWsheet2.Range("D54")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D55").Value = xlWorkSheet.Range("BJ" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("BJ" & nLigne)   'Si borne fontaine, préciser la source d'alimentation de la borne
                'xlDestRange = xlWsheet2.Range("D55")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D56").Value = xlWorkSheet.Range("BK" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("BK" & nLigne)   'Infrastructure fournissant l'eau construite par
                'xlDestRange = xlWsheet2.Range("D56")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D57").Value = xlWorkSheet.Range("BL" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("BL" & nLigne)   'Date de construction de l'infrastrucure fournissant l'eau
                'xlDestRange = xlWsheet2.Range("D57")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D58").Value = xlWorkSheet.Range("BM" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("BM" & nLigne)   'Etat général de l'infrastructure fournissant l'eau
                'xlDestRange = xlWsheet2.Range("D58")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D59").Value = xlWorkSheet.Range("BN" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("BN" & nLigne)   'Si l'eau est disponible sur le site de l'infrastructure, préciser la qualité de l'eau fournie
                'xlDestRange = xlWsheet2.Range("D59")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D60").Value = xlWorkSheet.Range("BO" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("BO" & nLigne)   'Si la qualité de l'eau est moyenne ou mauvaise, préciser la raison
                'xlDestRange = xlWsheet2.Range("D60")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D61").Value = xlWorkSheet.Range("BP" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("BP" & nLigne)   'Si autre, préciser
                'xlDestRange = xlWsheet2.Range("D61")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D62").Value = xlWorkSheet.Range("BQ" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("BQ" & nLigne)  'Source d'eau alternative (à préciser)
                'xlDestRange = xlWsheet2.Range("D62")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("J52").Value = xlWorkSheet.Range("BR" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("BR" & nLigne)   'Electricité disponible sur le site de l'établissement scolaire
                'xlDestRange = xlWsheet2.Range("J52")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("J53").Value = xlWorkSheet.Range("BS" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("BS" & nLigne)   'Infrastructure fournissant l'électricité
                'xlDestRange = xlWsheet2.Range("J53")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("J54").Value = xlWorkSheet.Range("BT" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("BT" & nLigne)   'Si autre, préciser la nature de l'infrastructure fournissant l'électricité
                'xlDestRange = xlWsheet2.Range("J54")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("J55").Value = xlWorkSheet.Range("BU" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("BU" & nLigne)    'Infrastructure fournissant l'électricité construite par
                'xlDestRange = xlWsheet2.Range("J55")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("J56").Value = xlWorkSheet.Range("BV" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("BV" & nLigne)    'Date de construction de l'infrastructure fournissant l'électricité
                'xlDestRange = xlWsheet2.Range("J56")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("J57").Value = xlWorkSheet.Range("BW" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("BW" & nLigne)   'Etat général de l'infrastructure fournissant l'électricité
                'xlDestRange = xlWsheet2.Range("J57")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne = 68

                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("BX" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("BX" & nLigne)   'Latrines disponibles  sur le site de l'établissement
                'xlDestRange = xlWsheet2.Range("D68")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("BY" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("BY" & nLigne)   'Nombre de latrines fonctionnelles
                'xlDestRange = xlWsheet2.Range("D69")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("BZ" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("BZ" & nLigne)  'Date de construction des latrines (année)
                'xlDestRange = xlWsheet2.Range("D70")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("CA" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("CA" & nLigne)   'Latrines construites par
                'xlDestRange = xlWsheet2.Range("D71")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("CB" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("CB" & nLigne)  'Etat général des latrines
                'xlDestRange = xlWsheet2.Range("D72")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("CC" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("CC" & nLigne)  'Remarques éventuelles concernant les latrines
                'xlDestRange = xlWsheet2.Range("D73")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne = 68

                xlWsheet2.Range("J" & SautLigne).Value = xlWorkSheet.Range("CD" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("CD" & nLigne)   'Douches sur le site de l'établissement
                'xlDestRange = xlWsheet2.Range("J68")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("J" & SautLigne).Value = xlWorkSheet.Range("CE" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("CE" & nLigne)   'Nombre de douches fonctionnelles
                'xlDestRange = xlWsheet2.Range("J69")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("J" & SautLigne).Value = xlWorkSheet.Range("CF" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("CF" & nLigne)   'Date de construction des douches (année)
                'xlDestRange = xlWsheet2.Range("J70")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("J" & SautLigne).Value = xlWorkSheet.Range("CG" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("CG" & nLigne)  'Douches construites par
                'xlDestRange = xlWsheet2.Range("J71")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("J" & SautLigne).Value = xlWorkSheet.Range("CH" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("CH" & nLigne)  'Etat général des douches
                'xlDestRange = xlWsheet2.Range("J72")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("J" & SautLigne).Value = xlWorkSheet.Range("CI" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("CI" & nLigne)   'Remarques éventuelles concernant les douches
                'xlDestRange = xlWsheet2.Range("J73")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("J" & SautLigne).Value = xlWorkSheet.Range("CJ" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("CJ" & nLigne)  'Nombre de bâtiments dans l'infrastructure (préciser)
                'xlDestRange = xlWsheet2.Range("H77")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne = 81
                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("CK" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("CK" & nLigne)   'Dimensions générales du batiment 1 (longueur x largeur)
                'xlDestRange = xlWsheet2.Range("D81")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("CL" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("CL" & nLigne)  'Etat général du bâtiment
                'xlDestRange = xlWsheet2.Range("D82")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("CM" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("CM" & nLigne)   'Type de sol
                'xlDestRange = xlWsheet2.Range("D83")
                'xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("CN" & nLigne)   'Type de sol/Plancher bois
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("CO" & nLigne)   'Type de sol/Dallage béton / Carrelage
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("CP" & nLigne)    'Type de sol/Plancher baobao
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("CQ" & nLigne)   'Type de sol/Terre battue
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("CR" & nLigne)   'Type de sol/Sable / gravier
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("CS" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("CS" & nLigne)   'Type de mur
                'xlDestRange = xlWsheet2.Range("D84")
                'xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("CT" & nLigne)   'Type de mur/Briques et enduit
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("CU" & nLigne)  'Type de mur/Agglos et enduit
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("CV" & nLigne)  'Type de mur/Planches bois
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("CW" & nLigne)   'Type de mur/Moellons de pierre et/ou enduit
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("CX" & nLigne)    'Type de mur/Baobao / Falafa
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("CY" & nLigne)   'Type de mur/Tole
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("CZ" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("CZ" & nLigne)   'Type de toiture
                'xlDestRange = xlWsheet2.Range("D85")
                'xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("DA" & nLigne)   'Type de toiture/Tuiles
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("DB" & nLigne)   'Type de toiture/Mokute
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("DC" & nLigne)   'Type de toiture/Falafa
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("DD" & nLigne)   'Type de toiture/Tole
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("DE" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("DE" & nLigne)   'Nombre de salles dans le bâtiment 1
                'xlDestRange = xlWsheet2.Range("D86")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("DF" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("DF" & nLigne)    'Affectation de la salle 1 (préciser)
                'xlDestRange = xlWsheet2.Range("D87")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("DG" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("DG" & nLigne)   'Affectation de la salle 2 (préciser)
                'xlDestRange = xlWsheet2.Range("D88")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("DH" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("DH" & nLigne)   'Affectation de la salle 3 (préciser)
                'xlDestRange = xlWsheet2.Range("D89")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("DI" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("DI" & nLigne)   'Affectation de la salle 4 (préciser)
                'xlDestRange = xlWsheet2.Range("D90")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("DJ" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("DJ" & nLigne)  'Affectation de la salle 5 (préciser)
                'xlDestRange = xlWsheet2.Range("D91")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("DK" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("DK" & nLigne)   'Affectation de la salle 6 (préciser)
                'xlDestRange = xlWsheet2.Range("D92")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("DL" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("DL" & nLigne)    'Affectation de la salle 7 (préciser)
                'xlDestRange = xlWsheet2.Range("D93")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("DM" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("DM" & nLigne)   'Affectation de la salle 8 (préciser)
                'xlDestRange = xlWsheet2.Range("D94")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne = 81
                xlWsheet2.Range("J" & SautLigne).Value = xlWorkSheet.Range("DN" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("DN" & nLigne)   'Dimensions générales du batiment 2 (longueur x largeur)
                'xlDestRange = xlWsheet2.Range("J81")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("J" & SautLigne).Value = xlWorkSheet.Range("DO" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("DO" & nLigne)   'Etat général du bâtiment
                'xlDestRange = xlWsheet2.Range("J82")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("H" & SautLigne).Value = xlWorkSheet.Range("DP" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("DP" & nLigne)   'Type de sol
                'xlDestRange = xlWsheet2.Range("H83")
                'xlSourceRange.Copy(xlDestRange)
                'xlWsheet2.Range("H83").Font.Size = 12

                ''xlSourceRange = xlWorkSheet.Range("DQ" & nLigne)    'Type de sol/Plancher bois
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("DR" & nLigne)   'Type de sol/Dallage béton / Carrelage
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("DS" & nLigne)   'Type de sol/Plancher baobao
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("DT" & nLigne)    'Type de sol/Terre battue
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("DU" & nLigne)   'Type de sol/Sable / gravier
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("H" & SautLigne).Value = xlWorkSheet.Range("DV" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("DV" & nLigne)   'Type de mur
                'xlDestRange = xlWsheet2.Range("H84")
                'xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("DW" & nLigne)    'Type de mur/Briques et enduit
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("DX" & nLigne)   'Type de mur/Agglos et enduit
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("DY" & nLigne)    'Type de mur/Planches bois
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)
                ''xlSourceRange = xlWorkSheet.Range("DZ" & nLigne)    'Type de mur/Moellons de pierre et/ou enduit
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("EA" & nLigne)  'Type de mur/Baobao / Falafa
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("EB" & nLigne)   'Type de mur/Tole
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("H" & SautLigne).Value = xlWorkSheet.Range("EC" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("EC" & nLigne)    'Type de toiture
                'xlDestRange = xlWsheet2.Range("H85")
                'xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("ED" & nLigne)   'Type de toiture/Tuiles
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("EE" & nLigne)    'Type de toiture/Mokute
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("EF" & nLigne)   'Type de toiture/Falafa
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("EG" & nLigne)    'Type de toiture/Tole
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("J" & SautLigne).Value = xlWorkSheet.Range("EH" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("EH" & nLigne)     'Nombre de salles dans le bâtiment 2
                'xlDestRange = xlWsheet2.Range("J86")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("J" & SautLigne).Value = xlWorkSheet.Range("EI" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("EI" & nLigne)   'Affectation de la salle 1 (préciser)
                'xlDestRange = xlWsheet2.Range("J87")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("J" & SautLigne).Value = xlWorkSheet.Range("EJ" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("EJ" & nLigne)   'Affectation de la salle 2 (préciser)
                'xlDestRange = xlWsheet2.Range("J88")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("J" & SautLigne).Value = xlWorkSheet.Range("EK" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("EK" & nLigne)   'Affectation de la salle 3 (préciser)
                'xlDestRange = xlWsheet2.Range("J89")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("J" & SautLigne).Value = xlWorkSheet.Range("EL" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("EL" & nLigne)   'Affectation de la salle 4 (préciser)
                'xlDestRange = xlWsheet2.Range("J90")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("J" & SautLigne).Value = xlWorkSheet.Range("EM" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("EM" & nLigne)   'Affectation de la salle 5 (préciser)
                'xlDestRange = xlWsheet2.Range("J91")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("J" & SautLigne).Value = xlWorkSheet.Range("EN" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("EN" & nLigne)   'Affectation de la salle 6 (préciser)
                'xlDestRange = xlWsheet2.Range("J92")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("J" & SautLigne).Value = xlWorkSheet.Range("EO" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("EO" & nLigne)   'Affectation de la salle 7 (préciser)
                'xlDestRange = xlWsheet2.Range("J93")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("J" & SautLigne).Value = xlWorkSheet.Range("EP" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("EP" & nLigne)   'Affectation de la salle 8 (préciser)
                'xlDestRange = xlWsheet2.Range("J94")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne = 98
                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("EQ" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("EQ" & nLigne)  'Dimensions générales du batiment 3 (longueur x largeur)
                'xlDestRange = xlWsheet2.Range("D98")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("ER" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("ER" & nLigne)   'Etat général du bâtiment
                'xlDestRange = xlWsheet2.Range("D99")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne = 100
                xlWsheet2.Range("B" & SautLigne).Value = xlWorkSheet.Range("ES" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("ES" & nLigne)   'Type de sol
                'xlDestRange = xlWsheet2.Range("B100")
                'xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("ET" & nLigne)  'Type de sol/Plancher bois
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("EU" & nLigne)   'Type de sol/Dallage béton / Carrelage
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("EV" & nLigne)    'Type de sol/Plancher baobao
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("EW" & nLigne)    'Type de sol/Terre battue
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("EX" & nLigne)   'Type de sol/Sable / gravier
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)
                ''xlWsheet2.Range("D15").Font.Size = 12

                SautLigne += 1
                xlWsheet2.Range("B" & SautLigne).Value = xlWorkSheet.Range("EY" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("EY" & nLigne)   'Type de mur
                'xlDestRange = xlWsheet2.Range("B101")
                'xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("EZ" & nLigne)   'Type de mur/Briques et enduit
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("FA" & nLigne)   'Type de mur/Agglos et enduit
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("FB" & nLigne)   'Type de mur/Planches bois
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("FC" & nLigne)    'Type de mur/Moellons de pierre et/ou enduit
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("FD" & nLigne)   'Type de mur/Baobao / Falafa
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("FE" & nLigne)    'Type de mur/Tole
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("B" & SautLigne).Value = xlWorkSheet.Range("FF" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("FF" & nLigne)    'Type de toiture
                'xlDestRange = xlWsheet2.Range("B102")
                'xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("FG" & nLigne)    'Type de toiture/Tuiles
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("FH" & nLigne)   'Type de toiture/Mokute
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("FI" & nLigne)   'Type de toiture/Falafa
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("FJ" & nLigne)    'Type de toiture/Tole
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("FK" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("FK" & nLigne)   'Nombre de salles dans le bâtiment 3
                'xlDestRange = xlWsheet2.Range("D103")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("FL" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("FL" & nLigne)   'Affectation de la salle 1 (préciser)
                'xlDestRange = xlWsheet2.Range("D104")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("FM" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("FM" & nLigne)  'Affectation de la salle 2 (préciser)
                'xlDestRange = xlWsheet2.Range("D105")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("FN" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("FN" & nLigne)   'Affectation de la salle 3 (préciser)
                'xlDestRange = xlWsheet2.Range("D106")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("FO" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("FO" & nLigne)   'Affectation de la salle 4 (préciser)
                'xlDestRange = xlWsheet2.Range("D107")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("FP" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("FP" & nLigne)   'Affectation de la salle 5 (préciser)
                'xlDestRange = xlWsheet2.Range("D108")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("FQ" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("FQ" & nLigne)   'Affectation de la salle 6 (préciser)
                'xlDestRange = xlWsheet2.Range("D109")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("FR" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("FR" & nLigne)   'Affectation de la salle 7 (préciser)
                'xlDestRange = xlWsheet2.Range("D110")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("FS" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("FS" & nLigne)   'Affectation de la salle 8 (préciser)
                'xlDestRange = xlWsheet2.Range("D111")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne = 98
                xlWsheet2.Range("J" & SautLigne).Value = xlWorkSheet.Range("FT" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("FT" & nLigne)  'Dimensions générales du batiment 4 (longueur x largeur)
                'xlDestRange = xlWsheet2.Range("J98")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("J" & SautLigne).Value = xlWorkSheet.Range("FU" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("FU" & nLigne)   'Etat général du bâtiment
                'xlDestRange = xlWsheet2.Range("J99")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("H" & SautLigne).Value = xlWorkSheet.Range("FV" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("FV" & nLigne)   'Type de sol
                'xlDestRange = xlWsheet2.Range("H100")
                'xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("FW" & nLigne)   'Type de sol/Plancher bois
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("FX" & nLigne)    'Type de sol/Dallage béton / Carrelage
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("FY" & nLigne)   'Type de sol/Plancher baobao
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("FZ" & nLigne)   'Type de sol/Terre battue
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("GA" & nLigne)   'Type de sol/Sable / gravier
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("H" & SautLigne).Value = xlWorkSheet.Range("GB" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("GB" & nLigne)   'Type de mur
                'xlDestRange = xlWsheet2.Range("H101")
                'xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("GC" & nLigne)   'Type de mur/Briques et enduit
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("GD" & nLigne)   'Type de mur/Agglos et enduit
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("GE" & nLigne)    'Type de mur/Planches bois
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("GF" & nLigne)   'Type de mur/Moellons de pierre et/ou enduit
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("GG" & nLigne)   'Type de mur/Baobao / Falafa
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("GH" & nLigne)   'Type de mur/Tole
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("GI" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("GI" & nLigne)    'Type de toiture
                'xlDestRange = xlWsheet2.Range("D102")
                'xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("GJ" & nLigne)  'Type de toiture/Tuiles
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("GK" & nLigne)   'Type de toiture/Mokute
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("GL" & nLigne)   'Type de toiture/Falafa
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("GM" & nLigne)    'Type de toiture/Tole
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("J" & SautLigne).Value = xlWorkSheet.Range("GN" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("GN" & nLigne)     'Nombre de salles dans le bâtiment 4
                'xlDestRange = xlWsheet2.Range("J103")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("J" & SautLigne).Value = xlWorkSheet.Range("GO" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("GO" & nLigne)    'Affectation de la salle 1 (préciser)
                'xlDestRange = xlWsheet2.Range("J104")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("J" & SautLigne).Value = xlWorkSheet.Range("GP" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("GP" & nLigne)    'Affectation de la salle 2 (préciser)
                'xlDestRange = xlWsheet2.Range("J105")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("J" & SautLigne).Value = xlWorkSheet.Range("GQ" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("GQ" & nLigne)   'Affectation de la salle 3 (préciser)
                'xlDestRange = xlWsheet2.Range("J106")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("J" & SautLigne).Value = xlWorkSheet.Range("GR" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("GR" & nLigne)    'Affectation de la salle 4 (préciser)
                'xlDestRange = xlWsheet2.Range("J107")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("J" & SautLigne).Value = xlWorkSheet.Range("GS" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("GS" & nLigne)   'Affectation de la salle 5 (préciser)
                'xlDestRange = xlWsheet2.Range("J108")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("J" & SautLigne).Value = xlWorkSheet.Range("GT" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("GT" & nLigne)    'Affectation de la salle 6 (préciser)
                'xlDestRange = xlWsheet2.Range("J109")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("J" & SautLigne).Value = xlWorkSheet.Range("GU" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("GU" & nLigne)    'Affectation de la salle 7 (préciser)
                'xlDestRange = xlWsheet2.Range("J110")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("J" & SautLigne).Value = xlWorkSheet.Range("GV" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("GV" & nLigne)   'Affectation de la salle 8 (préciser)
                'xlDestRange = xlWsheet2.Range("J111")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne = 118
                xlWsheet2.Range("C" & SautLigne).Value = xlWorkSheet.Range("GW" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("GW" & nLigne)   'Nombre total de tables banc
                'xlDestRange = xlWsheet2.Range("C118")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("GX" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("GX" & nLigne)   'Nombre en bon état
                'xlDestRange = xlWsheet2.Range("D118")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("E" & SautLigne).Value = xlWorkSheet.Range("GY" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("GY" & nLigne)   'Nombre en état d'usage
                'xlDestRange = xlWsheet2.Range("E118")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("F" & SautLigne).Value = xlWorkSheet.Range("GZ" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("GZ" & nLigne)   'Nombre en état vétuste mais fonctionnel
                'xlDestRange = xlWsheet2.Range("F118")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("H" & SautLigne).Value = xlWorkSheet.Range("HA" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("HA" & nLigne)   'Nombre inutilisable
                'xlDestRange = xlWsheet2.Range("H118")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("C" & SautLigne).Value = xlWorkSheet.Range("HB" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("HB" & nLigne)   'Nombre total de tableaux noirs
                'xlDestRange = xlWsheet2.Range("C119")
                'xlSourceRange.Copy(xlDestRange)
                'xlWsheet2.Range("C119").Font.Size = 12

                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("HC" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("HC" & nLigne)   'Nombre en bon état
                'xlDestRange = xlWsheet2.Range("D119")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("E" & SautLigne).Value = xlWorkSheet.Range("HD" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("HD" & nLigne)   'Nombre en état d'usage
                'xlDestRange = xlWsheet2.Range("E119")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("F" & SautLigne).Value = xlWorkSheet.Range("HE" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("HE" & nLigne)   'Nombre en état vétuste mais fonctionnel
                'xlDestRange = xlWsheet2.Range("F119")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("H" & SautLigne).Value = xlWorkSheet.Range("HF" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("HF" & nLigne)    'Nombre inutilisable
                'xlDestRange = xlWsheet2.Range("H119")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("C" & SautLigne).Value = xlWorkSheet.Range("HG" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("HG" & nLigne)   'Nombre total de tables de bureau
                'xlDestRange = xlWsheet2.Range("C120")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("HH" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("HH" & nLigne)   'Nombre en bon état
                'xlDestRange = xlWsheet2.Range("D120")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("E" & SautLigne).Value = xlWorkSheet.Range("HI" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("HI" & nLigne)   'Nombre en état d'usage
                'xlDestRange = xlWsheet2.Range("E120")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("F" & SautLigne).Value = xlWorkSheet.Range("HJ" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("HJ" & nLigne)   'Nombre en état vétuste mais fonctionnel
                'xlDestRange = xlWsheet2.Range("F120")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("H" & SautLigne).Value = xlWorkSheet.Range("HK" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("HK" & nLigne)   'Nombre inutilisable
                'xlDestRange = xlWsheet2.Range("H120")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("C" & SautLigne).Value = xlWorkSheet.Range("HL" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("HL" & nLigne)    'Nombre total de chaises
                'xlDestRange = xlWsheet2.Range("C121")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("HM" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("HM" & nLigne)   'Nombre en bon état
                'xlDestRange = xlWsheet2.Range("D121")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("E" & SautLigne).Value = xlWorkSheet.Range("HN" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("HN" & nLigne)    'Nombre en état d'usage
                'xlDestRange = xlWsheet2.Range("E121")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("F" & SautLigne).Value = xlWorkSheet.Range("HO" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("HO" & nLigne)   'Nombre en état vétuste mais fonctionnel
                'xlDestRange = xlWsheet2.Range("F121")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("H" & SautLigne).Value = xlWorkSheet.Range("HP" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("HP" & nLigne)   'Nombre inutilisable
                'xlDestRange = xlWsheet2.Range("H121")
                'xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("HQ" & nLigne)   'Y-a-t-il d'autres types de mobilier ?
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("A" & SautLigne).Value = xlWorkSheet.Range("HR" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("HR" & nLigne)   'Type de mobilier 1 (préciser)
                'xlDestRange = xlWsheet2.Range("A122")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("C" & SautLigne).Value = xlWorkSheet.Range("HS" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("HS" & nLigne)   'Nombre total de mobilier 1
                'xlDestRange = xlWsheet2.Range("C122")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("HT" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("HT" & nLigne)   'Nombre en bon état
                'xlDestRange = xlWsheet2.Range("D122")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("A" & SautLigne).Value = xlWorkSheet.Range("HU" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("HU" & nLigne)   'Nombre en état d'usage
                'xlDestRange = xlWsheet2.Range("E122")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("F" & SautLigne).Value = xlWorkSheet.Range("HV" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("HV" & nLigne)   'Nombre en état vétuste mais fonctionnel
                'xlDestRange = xlWsheet2.Range("F122")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("H" & SautLigne).Value = xlWorkSheet.Range("HW" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("HW" & nLigne)   'Nombre inutilisable
                'xlDestRange = xlWsheet2.Range("H122")
                'xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("HX" & nLigne)   'Y-a-t-il d'autres types de mobilier (2) ?
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("A" & SautLigne).Value = xlWorkSheet.Range("HY" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("HY" & nLigne)   'Type de mobilier 2 (préciser)
                'xlDestRange = xlWsheet2.Range("A123")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("C" & SautLigne).Value = xlWorkSheet.Range("HZ" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("HZ" & nLigne)   'Nombre total de mobilier 2
                'xlDestRange = xlWsheet2.Range("C123")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("IA" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("IA" & nLigne)   'Nombre en bon état
                'xlDestRange = xlWsheet2.Range("D123")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("E" & SautLigne).Value = xlWorkSheet.Range("IB" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("IB" & nLigne)  'Nombre en état d'usage
                'xlDestRange = xlWsheet2.Range("E123")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("F" & SautLigne).Value = xlWorkSheet.Range("IC" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("IC" & nLigne)   'Nombre en état vétuste mais fonctionnel
                'xlDestRange = xlWsheet2.Range("F123")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("H" & SautLigne).Value = xlWorkSheet.Range("ID" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("ID" & nLigne)   'Nombre inutilisable
                'xlDestRange = xlWsheet2.Range("H123")
                'xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("IE" & nLigne)   'Y-a-t-il d'autres types de mobilier (3) ?
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("A" & SautLigne).Value = xlWorkSheet.Range("IF" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("IF" & nLigne)   'Type de mobilier 3 (préciser)
                'xlDestRange = xlWsheet2.Range("A124")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("C" & SautLigne).Value = xlWorkSheet.Range("IG" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("IG" & nLigne)  'Nombre total de mobilier 3
                'xlDestRange = xlWsheet2.Range("C124")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("IH" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("IH" & nLigne)   'Nombre en bon état
                'xlDestRange = xlWsheet2.Range("D124")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("E" & SautLigne).Value = xlWorkSheet.Range("II" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("II" & nLigne)  'Nombre en état d'usage
                'xlDestRange = xlWsheet2.Range("E124")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("F" & SautLigne).Value = xlWorkSheet.Range("IJ" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("IJ" & nLigne)   'Nombre en état vétuste mais fonctionnel
                'xlDestRange = xlWsheet2.Range("F124")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("H" & SautLigne).Value = xlWorkSheet.Range("IK" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("IK" & nLigne)   'Nombre inutilisable
                'xlDestRange = xlWsheet2.Range("H124")
                'xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("IL" & nLigne)   'Y-a-t-il d'autres types de mobilier (4) ?
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("A" & SautLigne).Value = xlWorkSheet.Range("IM" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("IM" & nLigne)  'Type de mobilier 4 (préciser)
                'xlDestRange = xlWsheet2.Range("A125")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("C" & SautLigne).Value = xlWorkSheet.Range("IN" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("IN" & nLigne)   'Nombre total de mobilier 4
                'xlDestRange = xlWsheet2.Range("C125")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("IO" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("IO" & nLigne)   'Nombre en bon état
                'xlDestRange = xlWsheet2.Range("D125")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("E" & SautLigne).Value = xlWorkSheet.Range("IP" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("IP" & nLigne)   'Nombre en état d'usage
                'xlDestRange = xlWsheet2.Range("E125")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("F" & SautLigne).Value = xlWorkSheet.Range("IQ" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("IQ" & nLigne)   'Nombre en état vétuste mais fonctionnel
                'xlDestRange = xlWsheet2.Range("F125")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("H125" & SautLigne).Value = xlWorkSheet.Range("IR" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("IR" & nLigne)   'Nombre inutilisable
                'xlDestRange = xlWsheet2.Range("H125")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne = 129
                xlWsheet2.Range("H" & SautLigne).Value = xlWorkSheet.Range("IS" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("IS" & nLigne)   'De manière générale, le matériel pédagogique de base (manuel scolaire, matériel de géométrie, tableau noir,...) est-il disponible, en quantité suffisante et en état 
                'xlDestRange = xlWsheet2.Range("H129")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne = 132
                xlWsheet2.Range("C" & SautLigne).Value = xlWorkSheet.Range("IT" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("IT" & nLigne)   'Nombre total de manuels scolaire
                'xlDestRange = xlWsheet2.Range("C132")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("IU" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("IU" & nLigne)   'Nombre en bon état
                'xlDestRange = xlWsheet2.Range("D132")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("E" & SautLigne).Value = xlWorkSheet.Range("IV" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("IV" & nLigne)   'Nombre en état d'usage
                'xlDestRange = xlWsheet2.Range("E132")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("F" & SautLigne).Value = xlWorkSheet.Range("IW" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("IW" & nLigne)    'Nombre en état vétuste mais fonctionnel
                'xlDestRange = xlWsheet2.Range("F132")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("H" & SautLigne).Value = xlWorkSheet.Range("IX" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("IX" & nLigne)    'Nombre inutilisable
                'xlDestRange = xlWsheet2.Range("H132")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("C" & SautLigne).Value = xlWorkSheet.Range("IY" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("IY" & nLigne)    'Nombre total de règles
                'xlDestRange = xlWsheet2.Range("C133")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("IZ" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("IZ" & nLigne)    'Nombre en bon état
                'xlDestRange = xlWsheet2.Range("D133")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("E" & SautLigne).Value = xlWorkSheet.Range("JA" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("JA" & nLigne)   'Nombre en état d'usage
                'xlDestRange = xlWsheet2.Range("E133")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("F" & SautLigne).Value = xlWorkSheet.Range("JB" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("JB" & nLigne)    'Nombre en état vétuste mais fonctionnel
                'xlDestRange = xlWsheet2.Range("F133")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("H" & SautLigne).Value = xlWorkSheet.Range("JC" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("JC" & nLigne)   'Nombre inutilisable
                'xlDestRange = xlWsheet2.Range("H133")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("C" & SautLigne).Value = xlWorkSheet.Range("JD" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("JD" & nLigne)   'Nombre total de compas
                'xlDestRange = xlWsheet2.Range("C134")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("JE" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("JE" & nLigne)   'Nombre en bon état
                'xlDestRange = xlWsheet2.Range("D134")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("E" & SautLigne).Value = xlWorkSheet.Range("JF" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("JF" & nLigne)   'Nombre en état d'usage
                'xlDestRange = xlWsheet2.Range("E134")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("F" & SautLigne).Value = xlWorkSheet.Range("JG" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("JG" & nLigne)    'Nombre en état vétuste mais fonctionnel
                'xlDestRange = xlWsheet2.Range("F134")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("H" & SautLigne).Value = xlWorkSheet.Range("JH" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("JH" & nLigne)     'Nombre inutilisable
                'xlDestRange = xlWsheet2.Range("H134")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("C" & SautLigne).Value = xlWorkSheet.Range("JI" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("JI" & nLigne)     'Nombre total d'équerres
                'xlDestRange = xlWsheet2.Range("C135")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("JJ" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("JJ" & nLigne)   'Nombre en bon état
                'xlDestRange = xlWsheet2.Range("D135")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("E" & SautLigne).Value = xlWorkSheet.Range("JK" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("JK" & nLigne)   'Nombre en état d'usage
                'xlDestRange = xlWsheet2.Range("E135")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("F" & SautLigne).Value = xlWorkSheet.Range("JL" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("JL" & nLigne)   'Nombre en état vétuste mais fonctionnel
                'xlDestRange = xlWsheet2.Range("F135")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("H" & SautLigne).Value = xlWorkSheet.Range("JM" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("JM" & nLigne)   'Nombre inutilisable
                'xlDestRange = xlWsheet2.Range("H135")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("C" & SautLigne).Value = xlWorkSheet.Range("JN" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("JN" & nLigne)   'Nombre total de rapporteurs
                'xlDestRange = xlWsheet2.Range("C136")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("JO" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("JO" & nLigne)   'Nombre en bon état
                'xlDestRange = xlWsheet2.Range("D136")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("E" & SautLigne).Value = xlWorkSheet.Range("JP" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("JP" & nLigne)   'Nombre en état d'usage
                'xlDestRange = xlWsheet2.Range("E136")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("F" & SautLigne).Value = xlWorkSheet.Range("JQ" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("JQ" & nLigne)   'Nombre en état vétuste mais fonctionnel
                'xlDestRange = xlWsheet2.Range("F136")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("H" & SautLigne).Value = xlWorkSheet.Range("JR" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("JR" & nLigne)    'Nombre inutilisable
                'xlDestRange = xlWsheet2.Range("H136")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("C" & SautLigne).Value = xlWorkSheet.Range("JS" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("JS" & nLigne)    'Nombre total de planches murale
                'xlDestRange = xlWsheet2.Range("C137")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("JT" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("JT" & nLigne)   'Nombre en bon état
                'xlDestRange = xlWsheet2.Range("D137")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("E" & SautLigne).Value = xlWorkSheet.Range("JU" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("JU" & nLigne)   'Nombre en état d'usage
                'xlDestRange = xlWsheet2.Range("E137")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("F" & SautLigne).Value = xlWorkSheet.Range("JV" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("JV" & nLigne)    'Nombre en état vétuste mais fonctionnel
                'xlDestRange = xlWsheet2.Range("F137")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("H" & SautLigne).Value = xlWorkSheet.Range("JW" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("JW" & nLigne)   'Nombre inutilisable
                'xlDestRange = xlWsheet2.Range("H137")
                'xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("JX" & nLigne)   'Y-a-t-il d'autre type de matériel pédagogique ?
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("A" & SautLigne).Value = xlWorkSheet.Range("JY" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("JY" & nLigne)    'Préciser le type de matériel pédagogique 1
                'xlDestRange = xlWsheet2.Range("A138")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("C" & SautLigne).Value = xlWorkSheet.Range("JZ" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("JZ" & nLigne)    'Nombre total de matériel pédagogique de type 1
                'xlDestRange = xlWsheet2.Range("C138")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("KA" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("KA" & nLigne)   'Nombre en bon état
                'xlDestRange = xlWsheet2.Range("D138")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("E" & SautLigne).Value = xlWorkSheet.Range("KB" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("KB" & nLigne)   'Nombre en état d'usage
                'xlDestRange = xlWsheet2.Range("E138")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("F" & SautLigne).Value = xlWorkSheet.Range("KC" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("KC" & nLigne)   'Nombre en état vétuste mais fonctionnel
                'xlDestRange = xlWsheet2.Range("F138")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("H" & SautLigne).Value = xlWorkSheet.Range("KD" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("KD" & nLigne)   'Nombre inutilisable
                'xlDestRange = xlWsheet2.Range("H138")
                'xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("KE" & nLigne)    'Y-a-t-il d'autre type de matériel pédagogique ?
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("A" & SautLigne).Value = xlWorkSheet.Range("KF" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("KF" & nLigne)    'Préciser le type de matériel pédagogique 2
                'xlDestRange = xlWsheet2.Range("A139")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("C" & SautLigne).Value = xlWorkSheet.Range("KG" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("KG" & nLigne)    'Nombre total de matériel pédagogique de type 3 ???????
                'xlDestRange = xlWsheet2.Range("C139")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("KH" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("KH" & nLigne)    'Nombre en bon état
                'xlDestRange = xlWsheet2.Range("D139")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("E" & SautLigne).Value = xlWorkSheet.Range("KI" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("KI" & nLigne)   'Nombre en état d'usage
                'xlDestRange = xlWsheet2.Range("E139")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("F" & SautLigne).Value = xlWorkSheet.Range("KJ" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("KJ" & nLigne)    'Nombre en état vétuste mais fonctionnel
                'xlDestRange = xlWsheet2.Range("F139")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("H" & SautLigne).Value = xlWorkSheet.Range("KK" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("KK" & nLigne)   'Nombre inutilisable
                'xlDestRange = xlWsheet2.Range("H139")
                'xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("KL" & nLigne)   'Y-at-il d'autre type de matériel pédagogique ?
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("A" & SautLigne).Value = xlWorkSheet.Range("KM" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("KM" & nLigne)    'Préciser le type de matériel pédagogique 3
                'xlDestRange = xlWsheet2.Range("A140")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("C" & SautLigne).Value = xlWorkSheet.Range("KN" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("KN" & nLigne)    'Nombre total de matériel pédagogique de type 3
                'xlDestRange = xlWsheet2.Range("C140")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("KO" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("KO" & nLigne)     'Nombre en bon état
                'xlDestRange = xlWsheet2.Range("D140")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("E" & SautLigne).Value = xlWorkSheet.Range("KP" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("KP" & nLigne)   'Nombre en état d'usage
                'xlDestRange = xlWsheet2.Range("E140")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("F" & SautLigne).Value = xlWorkSheet.Range("KQ" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("KQ" & nLigne)    'Nombre en état vétuste mais fonctionnel
                'xlDestRange = xlWsheet2.Range("F140")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("H" & SautLigne).Value = xlWorkSheet.Range("KR" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("KR" & nLigne)    'Nombre inutilisable
                'xlDestRange = xlWsheet2.Range("H140")
                'xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("KS" & nLigne)   'Y-a-t-il d'autre type de matériel pédagogique ?
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("A" & SautLigne).Value = xlWorkSheet.Range("KT" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("KT" & nLigne)    'Préciser le type de matériel pédagogique 4
                'xlDestRange = xlWsheet2.Range("A141")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("C" & SautLigne).Value = xlWorkSheet.Range("KU" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("KU" & nLigne)    'Nombre total de matériel pédagogique de type 4
                'xlDestRange = xlWsheet2.Range("C141")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("KV" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("KV" & nLigne)    'Nombre en bon état
                'xlDestRange = xlWsheet2.Range("D141")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("E" & SautLigne).Value = xlWorkSheet.Range("KW" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("KW" & nLigne)    'Nombre en état d'usage
                'xlDestRange = xlWsheet2.Range("E141")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("F" & SautLigne).Value = xlWorkSheet.Range("KX" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("KX" & nLigne)   'Nombre en état vétuste mais fonctionnel
                'xlDestRange = xlWsheet2.Range("F141")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("H" & SautLigne).Value = xlWorkSheet.Range("KY" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("KY" & nLigne)   'Nombre inutilisable
                'xlDestRange = xlWsheet2.Range("H141")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne = 147
                xlWsheet2.Range("B" & SautLigne).Value = xlWorkSheet.Range("KZ" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("KZ" & nLigne)   'Nombre total d'élèves fréquentant l'établissement scolaire (tous niveaux confondus) pour l'année 2017/2018
                'xlDestRange = xlWsheet2.Range("B147")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("LA" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("LA" & nLigne)   'Nombre d'élèves féminins
                'xlDestRange = xlWsheet2.Range("D147")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("F" & SautLigne).Value = xlWorkSheet.Range("LB" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("LB" & nLigne)   'Nombre d'élèves masculins
                'xlDestRange = xlWsheet2.Range("F147")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne = 152
                xlWsheet2.Range("B" & SautLigne).Value = xlWorkSheet.Range("LC" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("LC" & nLigne)   'Nombre d'élèves en 11éme (année 2017/2018)
                'xlDestRange = xlWsheet2.Range("B152")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("C" & SautLigne).Value = xlWorkSheet.Range("LC" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("LD" & nLigne)   'Nombre d'élèves en 10ème (année 2017/2018)
                'xlDestRange = xlWsheet2.Range("C152")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("LE" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("LE" & nLigne)   'Nombre d'élèves en 9ème (année 2017/2018)
                'xlDestRange = xlWsheet2.Range("D152")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("E" & SautLigne).Value = xlWorkSheet.Range("LF" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("LF" & nLigne)   'Nombre d'élèves en 8ème (année 2017/2018)
                'xlDestRange = xlWsheet2.Range("E152")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("F" & SautLigne).Value = xlWorkSheet.Range("LG" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("LG" & nLigne)   'Nombre d'élèves en 7ème (année 2017/2018)
                'xlDestRange = xlWsheet2.Range("F152")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne = 157
                xlWsheet2.Range("B" & SautLigne).Value = xlWorkSheet.Range("LH" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("LH" & nLigne)   'Nombre d'élèves en 6ème (année 2017/2018)
                'xlDestRange = xlWsheet2.Range("B157")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("C" & SautLigne).Value = xlWorkSheet.Range("LI" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("LI" & nLigne)   'Nombre d'élèves en 5ème (année 2017/2018)
                'xlDestRange = xlWsheet2.Range("C157")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("LJ" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("LJ" & nLigne)   'Nombre d'élèves en 4ème (année 2017/2018)
                'xlDestRange = xlWsheet2.Range("D157")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("E" & SautLigne).Value = xlWorkSheet.Range("LF" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("LK" & nLigne)    'Nombre d'élèves en 3ème(année 2017/2018)
                'xlDestRange = xlWsheet2.Range("E157")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne = 162
                xlWsheet2.Range("B" & SautLigne).Value = xlWorkSheet.Range("LL" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("LL" & nLigne)   'Nombre d'élèves en 2nde (année 2017/2018)
                'xlDestRange = xlWsheet2.Range("B162")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("C" & SautLigne).Value = xlWorkSheet.Range("LM" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("LM" & nLigne)    'Nombre d'élèves en 1ère (année 2017/2018)
                'xlDestRange = xlWsheet2.Range("C162")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("LN" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("LN" & nLigne)    'Nombre d'élèves en terminale (année 2017/2018)
                'xlDestRange = xlWsheet2.Range("D162")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne = 148
                xlWsheet2.Range("B" & SautLigne).Value = xlWorkSheet.Range("LO" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("LO" & nLigne)   'Nombre total d'élèves fréquentant l'établissement scolaire (tous niveaux confondus) pour l'année 2016/2017
                'xlDestRange = xlWsheet2.Range("B148")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("LP" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("LP" & nLigne)   'Nombre d'élèves féminins
                'xlDestRange = xlWsheet2.Range("D148")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("F" & SautLigne).Value = xlWorkSheet.Range("LQ" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("LQ" & nLigne)   'Nombre d'élèves masculins
                'xlDestRange = xlWsheet2.Range("F148")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne = 153
                xlWsheet2.Range("B" & SautLigne).Value = xlWorkSheet.Range("LR" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("LR" & nLigne)  'Nombre d'élèves en 11éme (année 2016/2017)
                'xlDestRange = xlWsheet2.Range("B153")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("C" & SautLigne).Value = xlWorkSheet.Range("LS" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("LS" & nLigne)   'Nombre d'élèves en 10ème (année 2016/2017)
                'xlDestRange = xlWsheet2.Range("C153")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("LT" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("LT" & nLigne)   'Nombre d'élèves en 9ème (année 2016/2017)
                'xlDestRange = xlWsheet2.Range("D153")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("E" & SautLigne).Value = xlWorkSheet.Range("LU" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("LU" & nLigne)   'Nombre d'élèves en 8ème (année 2016/2017)
                'xlDestRange = xlWsheet2.Range("E153")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("F" & SautLigne).Value = xlWorkSheet.Range("LV" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("LV" & nLigne)   'Nombre d'élèves en 7ème (année 2016/2017)
                'xlDestRange = xlWsheet2.Range("F153")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne = 158
                xlWsheet2.Range("B" & SautLigne).Value = xlWorkSheet.Range("LW" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("LW" & nLigne)   'Nombre d'élèves en 6ème (année 2016/2017)
                'xlDestRange = xlWsheet2.Range("B158")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("C" & SautLigne).Value = xlWorkSheet.Range("LX" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("LX" & nLigne)  'Nombre d'élèves en 5ème (année 2016/2017)
                'xlDestRange = xlWsheet2.Range("C158")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("LY" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("LY" & nLigne)  'Nombre d'élèves en 4ème (année 2016/2017)
                'xlDestRange = xlWsheet2.Range("D158")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("E" & SautLigne).Value = xlWorkSheet.Range("LZ" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("LZ" & nLigne)   'Nombre d'élèves en 3ème (année 2016/2017)
                'xlDestRange = xlWsheet2.Range("E158")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne = 163
                xlWsheet2.Range("B" & SautLigne).Value = xlWorkSheet.Range("MA" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("MA" & nLigne)   'Nombre d'élèves en 2nde (année 2016/2017)
                'xlDestRange = xlWsheet2.Range("B163")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("C" & SautLigne).Value = xlWorkSheet.Range("MB" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("MB" & nLigne)   'Nombre d'élèves en 1ère (année 2016/2017)
                'xlDestRange = xlWsheet2.Range("C163")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("MC" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("MC" & nLigne)   'Nombre d'élèves en terminale (année 2016/2017)
                'xlDestRange = xlWsheet2.Range("D163")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne = 149
                xlWsheet2.Range("B" & SautLigne).Value = xlWorkSheet.Range("MD" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("MD" & nLigne)  'Nombre total d'élèves fréquentant l'établissement scolaire (tous niveaux confondus) pour l'année 2015/2016
                'xlDestRange = xlWsheet2.Range("B149")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("ME" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("ME" & nLigne)   'Nombre d'élèves féminins
                'xlDestRange = xlWsheet2.Range("D149")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("F" & SautLigne).Value = xlWorkSheet.Range("MF" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("MF" & nLigne)   'Nombre d'élèves masculins
                'xlDestRange = xlWsheet2.Range("F149")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne = 154
                xlWsheet2.Range("B" & SautLigne).Value = xlWorkSheet.Range("MG" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("MG" & nLigne)  'Nombre d'élèves en 11éme (année 2015/2016)
                'xlDestRange = xlWsheet2.Range("B154")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("C" & SautLigne).Value = xlWorkSheet.Range("MH" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("MH" & nLigne)  'Nombre d'élèves en 10ème (année 2015/2016)
                'xlDestRange = xlWsheet2.Range("C154")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("MI" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("MI" & nLigne)   'Nombre d'élèves en 9ème (année 2015/2016)
                'xlDestRange = xlWsheet2.Range("D154")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("E" & SautLigne).Value = xlWorkSheet.Range("MJ" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("MJ" & nLigne)   'Nombre d'élèves en 8ème (année 2015/2016)
                'xlDestRange = xlWsheet2.Range("E154")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("F" & SautLigne).Value = xlWorkSheet.Range("MK" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("MK" & nLigne)   'Nombre d'élèves en 7ème (année 2015/2016)
                'xlDestRange = xlWsheet2.Range("F154")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne = 159
                xlWsheet2.Range("B" & SautLigne).Value = xlWorkSheet.Range("ML" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("ML" & nLigne)   'Nombre d'élèves en 6ème (année 2015/2016)
                'xlDestRange = xlWsheet2.Range("B159")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("C" & SautLigne).Value = xlWorkSheet.Range("MM" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("MM" & nLigne)    'Nombre d'élèves en 5ème (année 2015/2016)
                'xlDestRange = xlWsheet2.Range("C159")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("MN" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("MN" & nLigne)    'Nombre d'élèves en 4ème (année 2015/2016)
                'xlDestRange = xlWsheet2.Range("D159")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("E" & SautLigne).Value = xlWorkSheet.Range("MO" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("MO" & nLigne)    'Nombre d'élèves en 3ème (année 2015/2016)
                'xlDestRange = xlWsheet2.Range("E159")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne = 164
                xlWsheet2.Range("B" & SautLigne).Value = xlWorkSheet.Range("MP" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("MP" & nLigne)   'Nombre d'élèves en 2nde (année 2015/2016)
                'xlDestRange = xlWsheet2.Range("B164")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("C" & SautLigne).Value = xlWorkSheet.Range("MQ" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("MQ" & nLigne)   'Nombre d'élèves en 1ère (année 2015/2016)
                'xlDestRange = xlWsheet2.Range("C164")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("MR" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("MR" & nLigne)   'Nombre d'élèves en terminale (année 2015/2016)
                'xlDestRange = xlWsheet2.Range("D164")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne = 169
                xlWsheet2.Range("B" & SautLigne).Value = xlWorkSheet.Range("MS" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("MS" & nLigne)   'Nombre total d'enseignants dans l'établissement
                'xlDestRange = xlWsheet2.Range("B169")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("C" & SautLigne).Value = xlWorkSheet.Range("MT" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("MT" & nLigne)    'Nombre d'enseignants fonctionnaires
                'xlDestRange = xlWsheet2.Range("C169")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("F" & SautLigne).Value = xlWorkSheet.Range("MU" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("MU" & nLigne)    'Nombre d'enseignants rémunérés par des subventions d'état
                'xlDestRange = xlWsheet2.Range("F169")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("H" & SautLigne).Value = xlWorkSheet.Range("MV" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("MV" & nLigne)    'Nombre d'enseignant rémunérés par FRAM
                'xlDestRange = xlWsheet2.Range("H169")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne = 173
                xlWsheet2.Range("H" & SautLigne).Value = xlWorkSheet.Range("MW" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("MW" & nLigne)    'Nombre total de personnels (hors enseignants) dans l'établissement
                'xlDestRange = xlWsheet2.Range("H173")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne = 176
                xlWsheet2.Range("A" & SautLigne).Value = xlWorkSheet.Range("MX" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("MX" & nLigne)    'Fonction du personnel 1 (préciser)
                'xlDestRange = xlWsheet2.Range("A176")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("B" & SautLigne).Value = xlWorkSheet.Range("MY" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("MY" & nLigne)   'Prise en charge du personnel 1 (préciser)
                'xlDestRange = xlWsheet2.Range("B176")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("A" & SautLigne).Value = xlWorkSheet.Range("MZ" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("MZ" & nLigne)    'Fonction du personnel 2 (préciser)
                'xlDestRange = xlWsheet2.Range("A177")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("B" & SautLigne).Value = xlWorkSheet.Range("NA" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("NA" & nLigne)    'Prise en charge du personnel 2 (préciser)
                'xlDestRange = xlWsheet2.Range("B177")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("A" & SautLigne).Value = xlWorkSheet.Range("NB" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("NB" & nLigne)    'Fonction du personnel 3 (préciser)
                'xlDestRange = xlWsheet2.Range("A178")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("B" & SautLigne).Value = xlWorkSheet.Range("NC" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("NC" & nLigne)    'Prise en charge du personnel 3 (préciser)
                'xlDestRange = xlWsheet2.Range("B178")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("A" & SautLigne).Value = xlWorkSheet.Range("ND" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("ND" & nLigne)     'Fonction du personnel 4 (préciser)
                'xlDestRange = xlWsheet2.Range("A179")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("B" & SautLigne).Value = xlWorkSheet.Range("NE" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("NE" & nLigne)     'Prise en charge du personnel 4 (préciser)
                'xlDestRange = xlWsheet2.Range("B179")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("A" & SautLigne).Value = xlWorkSheet.Range("NF" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("NF" & nLigne)      'Fonction du personnel 5 (préciser)
                'xlDestRange = xlWsheet2.Range("A180")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("B" & SautLigne).Value = xlWorkSheet.Range("NG" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("NG" & nLigne)     'Prise en charge du personnel 5 (préciser)
                'xlDestRange = xlWsheet2.Range("B180")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("A" & SautLigne).Value = xlWorkSheet.Range("NH" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("NH" & nLigne)      'Fonction du personnel 6 (préciser)
                'xlDestRange = xlWsheet2.Range("A181")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("B" & SautLigne).Value = xlWorkSheet.Range("NI" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("NI" & nLigne)      'Prise en charge du personnel 6 (préciser)
                'xlDestRange = xlWsheet2.Range("B181")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne = 189
                xlWsheet2.Range("A" & SautLigne).Value = xlWorkSheet.Range("NJ" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("NJ" & nLigne)   'Description de la maintenance effectuée (1)
                'xlDestRange = xlWsheet2.Range("A189")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("C" & SautLigne).Value = xlWorkSheet.Range("NK" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("NK" & nLigne)   'Raison de la maintenance effectuée (1) (et préciser si maintenance normale ou exceptionnelle)
                'xlDestRange = xlWsheet2.Range("C189")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("E" & SautLigne).Value = xlWorkSheet.Range("NL" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("NL" & nLigne)    'Date de la dernière maintenance effectuée (1)
                'xlDestRange = xlWsheet2.Range("E189")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("F" & SautLigne).Value = xlWorkSheet.Range("NM" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("NM" & nLigne)    'Maintenance effectuée (1) par
                'xlDestRange = xlWsheet2.Range("F189")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("G" & SautLigne).Value = xlWorkSheet.Range("NN" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("NN" & nLigne)    'Coût de la maintenance effectuée (1) en Ariary
                'xlDestRange = xlWsheet2.Range("G189")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("H" & SautLigne).Value = xlWorkSheet.Range("NO" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("NO" & nLigne)    'Origine du financement de la maintenance effectuée (1)
                'xlDestRange = xlWsheet2.Range("H189")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("I" & SautLigne).Value = xlWorkSheet.Range("NP" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("NP" & nLigne)   'Remarques éventuelles concernant la maintenance effectuée (1)
                'xlDestRange = xlWsheet2.Range("I189")
                'xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("NQ" & nLigne)   'Y-a-t-il une autre maintenance effectuée (2) à renseigner ?
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("A" & SautLigne).Value = xlWorkSheet.Range("NR" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("NR" & nLigne)   'Description de la maintenance effectuée (2)
                'xlDestRange = xlWsheet2.Range("A190")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("C" & SautLigne).Value = xlWorkSheet.Range("NS" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("NS" & nLigne)   'Raison de la maintenance effectuée (2) (et préciser si maintenance normale ou exceptionnelle)
                'xlDestRange = xlWsheet2.Range("C190")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("E" & SautLigne).Value = xlWorkSheet.Range("NT" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("NT" & nLigne)   'Date de la dernière maintenance effectuée (2)
                'xlDestRange = xlWsheet2.Range("E190")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("F" & SautLigne).Value = xlWorkSheet.Range("NU" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("NU" & nLigne)   'Maintenance effectuée (2) par
                'xlDestRange = xlWsheet2.Range("F190")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("G" & SautLigne).Value = xlWorkSheet.Range("NV" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("NV" & nLigne)   'Coût de la maintenance effectuée (2) en Ariary
                'xlDestRange = xlWsheet2.Range("G190")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("H" & SautLigne).Value = xlWorkSheet.Range("NW" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("NW" & nLigne)    'Origine du financement de la maintenance effectuée (2)
                'xlDestRange = xlWsheet2.Range("H190")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("I" & SautLigne).Value = xlWorkSheet.Range("NX" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("NX" & nLigne)   'Remarques éventuelles concernant la maintenance effectuée (2)
                'xlDestRange = xlWsheet2.Range("I190")
                'xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("NY" & nLigne)   'Y-a-t-il une autre maintenance effectuée (3) à renseigner ?
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("A" & SautLigne).Value = xlWorkSheet.Range("NZ" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("NZ" & nLigne)   'Description de la maintenance effectuée (3)
                'xlDestRange = xlWsheet2.Range("A191")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("C" & SautLigne).Value = xlWorkSheet.Range("OA" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("OA" & nLigne)    'Raison de la maintenance effectuée (3) (et préciser si maintenance normale ou exceptionnelle)
                'xlDestRange = xlWsheet2.Range("C191")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("E" & SautLigne).Value = xlWorkSheet.Range("OB" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("OB" & nLigne)   'Date de la dernière maintenance effectuée (3)
                'xlDestRange = xlWsheet2.Range("E191")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("F" & SautLigne).Value = xlWorkSheet.Range("OC" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("OC" & nLigne)   'Maintenance effectuée (3) par
                'xlDestRange = xlWsheet2.Range("F191")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("G" & SautLigne).Value = xlWorkSheet.Range("OD" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("OD" & nLigne)   'Coût de la maintenance effectuée (3) en Ariary
                'xlDestRange = xlWsheet2.Range("G191")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("H" & SautLigne).Value = xlWorkSheet.Range("OE" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("OE" & nLigne)   'Origine du financement de la maintenance effectuée (3)
                'xlDestRange = xlWsheet2.Range("H191")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("I" & SautLigne).Value = xlWorkSheet.Range("OF" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("OF" & nLigne)   'Remarques éventuelles concernant la maintenance effectuée (3)
                'xlDestRange = xlWsheet2.Range("I191")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne = 197
                xlWsheet2.Range("A" & SautLigne).Value = xlWorkSheet.Range("OG" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("OG" & nLigne)    'Description de la maintenance prévue (1)
                'xlDestRange = xlWsheet2.Range("A197")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("E" & SautLigne).Value = xlWorkSheet.Range("OH" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("OH" & nLigne)   'Date de la maintenance prévue (1)
                'xlDestRange = xlWsheet2.Range("E197")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("C" & SautLigne).Value = xlWorkSheet.Range("OI" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("OI" & nLigne)    'Raison de la maintenance prévue (1) (maintenance normale ou exceptionnelle)
                'xlDestRange = xlWsheet2.Range("C197")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("F" & SautLigne).Value = xlWorkSheet.Range("OJ" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("OJ" & nLigne)    'Maintenance prévue (1) d'être effectuée par
                'xlDestRange = xlWsheet2.Range("F197")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("G" & SautLigne).Value = xlWorkSheet.Range("OK" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("OK" & nLigne)    'Coût prévu pour la maintenance (1) en Ariary
                'xlDestRange = xlWsheet2.Range("G197")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("H" & SautLigne).Value = xlWorkSheet.Range("OL" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("OL" & nLigne)   'Origine du financement pour la maintenance prévue (1)
                'xlDestRange = xlWsheet2.Range("H197")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("I" & SautLigne).Value = xlWorkSheet.Range("OM" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("OM" & nLigne)    'Remarques éventuelles concernant la maintenance prévue (1)
                'xlDestRange = xlWsheet2.Range("I197")
                'xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("ON" & nLigne)   'Y-a-t-il une autre maintenance prévue (2) à renseigner ?
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("A" & SautLigne).Value = xlWorkSheet.Range("OO" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("OO" & nLigne)   'Description de la maintenance prévue (2)
                'xlDestRange = xlWsheet2.Range("A198")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("E" & SautLigne).Value = xlWorkSheet.Range("OP" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("OP" & nLigne)   'Date de la maintenance prévue (2)
                'xlDestRange = xlWsheet2.Range("E198")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("C" & SautLigne).Value = xlWorkSheet.Range("OQ" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("OQ" & nLigne)   'Raison de la maintenance prévue (2) (maintenance normale ou exceptionnelle)
                'xlDestRange = xlWsheet2.Range("C198")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("F" & SautLigne).Value = xlWorkSheet.Range("OR" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("OR" & nLigne)    'Maintenance prévue (2) d'être effectuée par
                'xlDestRange = xlWsheet2.Range("F198")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("G" & SautLigne).Value = xlWorkSheet.Range("OS" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("OS" & nLigne)    'Coût prévu pour la maintenance (2) en Ariary
                'xlDestRange = xlWsheet2.Range("G198")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("H" & SautLigne).Value = xlWorkSheet.Range("OT" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("OT" & nLigne)    'Origine du financement pour la maintenance prévue (2)
                'xlDestRange = xlWsheet2.Range("H198")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("I" & SautLigne).Value = xlWorkSheet.Range("OU" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("OU" & nLigne)   'Remarques éventuelles concernant la maintenance prévue (2)
                'xlDestRange = xlWsheet2.Range("I198")
                'xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("OV" & nLigne)   'Y-a-t-il une autre maintenance prévue (3) à renseigner ?
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("A" & SautLigne).Value = xlWorkSheet.Range("OW" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("OW" & nLigne)    'Description de la maintenance prévue (3)
                'xlDestRange = xlWsheet2.Range("A199")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("E" & SautLigne).Value = xlWorkSheet.Range("OX" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("OX" & nLigne)    'Date de la maintenance prévue (3)
                'xlDestRange = xlWsheet2.Range("E199")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("C" & SautLigne).Value = xlWorkSheet.Range("OY" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("OY" & nLigne)    'Raison de la maintenance prévue (3) (maintenance normale ou exceptionnelle)
                'xlDestRange = xlWsheet2.Range("C199")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("F" & SautLigne).Value = xlWorkSheet.Range("OZ" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("OZ" & nLigne)   'Maintenance prévue (3) d'être effectuée par
                'xlDestRange = xlWsheet2.Range("F199")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("G" & SautLigne).Value = xlWorkSheet.Range("PA" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("PA" & nLigne)    'Coût prévu pour la maintenance (3) en Ariary
                'xlDestRange = xlWsheet2.Range("G199")
                'xlSourceRange.Copy(xlDestRange)
                'xlWsheet2.Range("G199").Font.Size = 12

                xlWsheet2.Range("H" & SautLigne).Value = xlWorkSheet.Range("PB" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("PB" & nLigne)    'Origine du financement pour la maintenance prévue (3)
                'xlDestRange = xlWsheet2.Range("H199")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("I" & SautLigne).Value = xlWorkSheet.Range("PC" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("PC" & nLigne)   'Remarques éventuelles concernant la maintenance prévue (3)
                'xlDestRange = xlWsheet2.Range("I199")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne = 206
                xlWsheet2.Range("F" & SautLigne).Value = xlWorkSheet.Range("PD" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("PD" & nLigne)   'Préciser le mode de gestion de l'infrastructure
                'xlDestRange = xlWsheet2.Range("F206")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("E" & SautLigne).Value = xlWorkSheet.Range("PE" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("PE" & nLigne)  'Si autre, préciser le mode de gestion
                'xlDestRange = xlWsheet2.Range("E207")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne = 212
                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("PF" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("PF" & nLigne)   'Nom du Directeur/trice
                'xlDestRange = xlWsheet2.Range("D212")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("F" & SautLigne).Value = xlWorkSheet.Range("PG" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("PG" & nLigne)   'Téléphone du Directeur/trice
                'xlDestRange = xlWsheet2.Range("F212")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("H" & SautLigne).Value = xlWorkSheet.Range("PH" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("PH" & nLigne)   'Nombre d'années d'expérience dans le poste
                'xlDestRange = xlWsheet2.Range("H212")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("PI" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("PI" & nLigne)   'Nom du Président/e du FRAM
                'xlDestRange = xlWsheet2.Range("D213")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("F" & SautLigne).Value = xlWorkSheet.Range("PJ" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("PJ" & nLigne)   'Téléphone du Président/e du FRAM
                'xlDestRange = xlWsheet2.Range("F213")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("H" & SautLigne).Value = xlWorkSheet.Range("PK" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("PK" & nLigne)    'Nombre d'années d'expérience dans le poste
                'xlDestRange = xlWsheet2.Range("H213")
                'xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("PL" & nLigne)   'Y-a-t-il d'autre personnel
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("A" & SautLigne).Value = xlWorkSheet.Range("PM" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("PM" & nLigne)   'Fonction du personnel (1)
                'xlDestRange = xlWsheet2.Range("A214")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("PN" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("PN" & nLigne)   'Nom du personnel (1)
                'xlDestRange = xlWsheet2.Range("D214")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("F" & SautLigne).Value = xlWorkSheet.Range("PO" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("PO" & nLigne)   'Téléphone du personnel (1)
                'xlDestRange = xlWsheet2.Range("F214")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("H" & SautLigne).Value = xlWorkSheet.Range("PP" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("PP" & nLigne)    'Nombre d'années d'expérience dans le poste
                'xlDestRange = xlWsheet2.Range("H214")
                'xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("PQ" & nLigne)   'Y-a-t-il un autre personnel (2) à renseigner ?
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("A" & SautLigne).Value = xlWorkSheet.Range("PR" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("PR" & nLigne)   'Fonction du personnel (2)
                'xlDestRange = xlWsheet2.Range("A215")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("PS" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("PS" & nLigne)  'Nom du personnel (2)
                'xlDestRange = xlWsheet2.Range("D215")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("F" & SautLigne).Value = xlWorkSheet.Range("PT" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("PT" & nLigne)   'Téléphone du personnel (2)
                'xlDestRange = xlWsheet2.Range("F215")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("H" & SautLigne).Value = xlWorkSheet.Range("PU" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("PU" & nLigne)   'Nombre d'années d'expérience dans le poste
                'xlDestRange = xlWsheet2.Range("H215")
                'xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("PV" & nLigne)   'Y-a-t-il un autre personnel (3) à renseigner ?
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                SautLigne = 216
                xlWsheet2.Range("A" & SautLigne).Value = xlWorkSheet.Range("PW" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("PW" & nLigne)   'Fonction du personnel (3)
                'xlDestRange = xlWsheet2.Range("A216")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("PX" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("PX" & nLigne)   'Nom du personnel (3)
                'xlDestRange = xlWsheet2.Range("D216")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("F" & SautLigne).Value = xlWorkSheet.Range("PY" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("PY" & nLigne)   'Téléphone du personnel (3)
                'xlDestRange = xlWsheet2.Range("F216")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("H" & SautLigne).Value = xlWorkSheet.Range("PZ" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("PZ" & nLigne)   'Nombre d'années d'expérience dans le poste
                'xlDestRange = xlWsheet2.Range("H216")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne = 221
                xlWsheet2.Range("A" & SautLigne).Value = xlWorkSheet.Range("QA" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("QA" & nLigne)   'Objet de la dernière réunion (1) (préciser)
                'xlDestRange = xlWsheet2.Range("A221")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("B" & SautLigne).Value = xlWorkSheet.Range("QB" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("QB" & nLigne)   'Dernière réunion (1) régulière ?
                'xlDestRange = xlWsheet2.Range("B221")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("C" & SautLigne).Value = xlWorkSheet.Range("QC" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("QC" & nLigne)   'Date de cette réunion (1)
                'xlDestRange = xlWsheet2.Range("C221")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("E" & SautLigne).Value = xlWorkSheet.Range("QD" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("QD" & nLigne)   'La commune a-t-elle été informée de cette réunion (1) ?
                'xlDestRange = xlWsheet2.Range("E221")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("F" & SautLigne).Value = xlWorkSheet.Range("QE" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("QE" & nLigne)    'Emission d'un PV pour la dernière réunion (1) ?
                'xlDestRange = xlWsheet2.Range("F221")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("G" & SautLigne).Value = xlWorkSheet.Range("QF" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("QF" & nLigne)   'Participants à la dernière réunion (1)
                'xlDestRange = xlWsheet2.Range("G221")
                'xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("QG" & nLigne)   'Participants à la dernière réunion (1)/Directeur/trice
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("QH" & nLigne).Text
                ''xlSourceRange = xlWorkSheet.Range("QH" & nLigne)   'Participants à la dernière réunion (1)/Président/e FRAM
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("QI" & nLigne)   'Participants à la dernière réunion (1)/Enseignants
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("QJ" & nLigne)  'Participants à la dernière réunion (1)/Parents d'élèves
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("I" & SautLigne).Value = xlWorkSheet.Range("QK" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("QK" & nLigne)   'Remarques particulières concernant cette dernière réunion (1)
                'xlDestRange = xlWsheet2.Range("I221")
                'xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("QL" & nLigne)   'Y-a-t-il une autre réunion à renseigner ?
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("A" & SautLigne).Value = xlWorkSheet.Range("QM" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("QM" & nLigne)   'Objet de la dernière réunion (2) (préciser)
                'xlDestRange = xlWsheet2.Range("A222")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("B" & SautLigne).Value = xlWorkSheet.Range("QN" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("QN" & nLigne)    'Dernière réunion (2) régulière ?
                'xlDestRange = xlWsheet2.Range("B222")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("C" & SautLigne).Value = xlWorkSheet.Range("QO" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("QO" & nLigne)   'Date de cette réunion (2)
                'xlDestRange = xlWsheet2.Range("C222")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("QP" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("QP" & nLigne)   'La commune a-t-elle été informée de cette réunion (2) ?
                'xlDestRange = xlWsheet2.Range("D222")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("F" & SautLigne).Value = xlWorkSheet.Range("QQ" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("QQ" & nLigne)   'Emission d'un PV pour la dernière réunion (2) ?
                'xlDestRange = xlWsheet2.Range("F222")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("G" & SautLigne).Value = xlWorkSheet.Range("QR" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("QR" & nLigne)   'Participants à la dernière réunion (2)
                'xlDestRange = xlWsheet2.Range("G222")
                'xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("QS" & nLigne)   'Participants à la dernière réunion (2)/Directeur/trice
                ''xlDestRange = xlWsheet2.Range("I222")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("QT" & nLigne)   'Participants à la dernière réunion (2)/Président/e FRAM
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("QU" & nLigne)     'Participants à la dernière réunion (2)/Enseignants
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("QV" & nLigne)   'Participants à la dernière réunion (2)/Parents d'élèves
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("I" & SautLigne).Value = xlWorkSheet.Range("QW" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("QW" & nLigne)   'Remarques particulières concernant cette dernière réunion (2)
                'xlDestRange = xlWsheet2.Range("I222")
                'xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("QX" & nLigne)   'Y-a-t-il une autre réunion à renseigner ?
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("A" & SautLigne).Value = xlWorkSheet.Range("QY" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("QY" & nLigne)   'Objet de la dernière réunion (3) (préciser)
                'xlDestRange = xlWsheet2.Range("A223")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("B" & SautLigne).Value = xlWorkSheet.Range("QZ" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("QZ" & nLigne)     'Dernière réunion (3) régulière ?
                'xlDestRange = xlWsheet2.Range("B223")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("C" & SautLigne).Value = xlWorkSheet.Range("RA" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("RA" & nLigne)   'Date de cette réunion (3)
                'xlDestRange = xlWsheet2.Range("C223")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("RB" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("RB" & nLigne)   'La commune a-t-elle été informée de cette réunion (3) ?
                'xlDestRange = xlWsheet2.Range("D223")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("F" & SautLigne).Value = xlWorkSheet.Range("RC" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("RC" & nLigne)   'Emission d'un PV pour la dernière réunion (3) ?
                'xlDestRange = xlWsheet2.Range("F223")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("G" & SautLigne).Value = xlWorkSheet.Range("RD" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("RD" & nLigne)    'Participants à la dernière réunion (1)
                'xlDestRange = xlWsheet2.Range("G223")
                'xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("RE" & nLigne)   'Participants à la dernière réunion (1)/Directeur/trice
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("RF" & nLigne)    'Participants à la dernière réunion (1)/Président/e FRAM
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("RG" & nLigne)    'Participants à la dernière réunion (1)/Enseignants
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("RH" & nLigne)   'Participants à la dernière réunion (1)/Parents d'élèves
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("I" & SautLigne).Value = xlWorkSheet.Range("RI" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("RI" & nLigne)   'Remarques particulières concernant cette dernière réunion (1)
                'xlDestRange = xlWsheet2.Range("I223")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne = 228
                xlWsheet2.Range("A" & SautLigne).Value = xlWorkSheet.Range("RJ" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("RJ" & nLigne)   'Origine du financement 1 (préciser)
                'xlDestRange = xlWsheet2.Range("A228")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("C" & SautLigne).Value = xlWorkSheet.Range("RK" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("RK" & nLigne)    'Montant du financement 1 (en Ariary)
                'xlDestRange = xlWsheet2.Range("C228")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("RL" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("RL" & nLigne)   'Affectation du financement 1
                'xlDestRange = xlWsheet2.Range("D228")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("E" & SautLigne).Value = xlWorkSheet.Range("RM" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("RM" & nLigne)    'Date / périodicité du financement 1
                'xlDestRange = xlWsheet2.Range("E228")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("G" & SautLigne).Value = xlWorkSheet.Range("RN" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("RN" & nLigne)    'Administration / gestion des fonds 1 (qui gère ?)
                'xlDestRange = xlWsheet2.Range("G228")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("H" & SautLigne).Value = xlWorkSheet.Range("RO" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("RO" & nLigne)    'Remarques particulières concernant le financement 1
                'xlDestRange = xlWsheet2.Range("H228")
                'xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("RP" & nLigne)   'Y-a-t-il un autre financement (2) ?
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("A" & SautLigne).Value = xlWorkSheet.Range("RQ" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("RQ" & nLigne)   'Origine du financement 2 (préciser)
                'xlDestRange = xlWsheet2.Range("A229")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("C" & SautLigne).Value = xlWorkSheet.Range("RR" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("RR" & nLigne)   'Montant du financement 2 (en Ariary)
                'xlDestRange = xlWsheet2.Range("C229")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("RS" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("RS" & nLigne)   'Affectation du financement 2
                'xlDestRange = xlWsheet2.Range("D229")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("E" & SautLigne).Value = xlWorkSheet.Range("RT" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("RT" & nLigne)   'Date / périodicité du financement 2
                'xlDestRange = xlWsheet2.Range("E229")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("G" & SautLigne).Value = xlWorkSheet.Range("RU" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("RU" & nLigne)   'Administration / gestion des fonds 2 (qui gère ?)
                'xlDestRange = xlWsheet2.Range("G229")
                'xlSourceRange.Copy(xlDestRange)
                'xlWsheet2.Range("G229").Font.Size = 12

                xlWsheet2.Range("H" & SautLigne).Value = xlWorkSheet.Range("RV" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("RV" & nLigne)   'Remarques particulières concernant le financement 2
                'xlDestRange = xlWsheet2.Range("H229")
                'xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("RW" & nLigne)   'Y-a-t-il un autre financement (3) ?
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("A" & SautLigne).Value = xlWorkSheet.Range("RX" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("RX" & nLigne)   'Origine du financement 3 (préciser)
                'xlDestRange = xlWsheet2.Range("A230")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("C" & SautLigne).Value = xlWorkSheet.Range("RY" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("RY" & nLigne)   'Montant du financement 3 (en Ariary)
                'xlDestRange = xlWsheet2.Range("C230")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("RZ" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("RZ" & nLigne)    'Affectation du financement 3
                'xlDestRange = xlWsheet2.Range("D230")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("E" & SautLigne).Value = xlWorkSheet.Range("SA" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("SA" & nLigne)   'Date / périodicité du financement 3 
                'xlDestRange = xlWsheet2.Range("E230")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("G" & SautLigne).Value = xlWorkSheet.Range("SB" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("SB" & nLigne)   'Administration / gestion des fonds 3 (qui gère ?)
                'xlDestRange = xlWsheet2.Range("G230")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("H" & SautLigne).Value = xlWorkSheet.Range("SC" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("SC" & nLigne)    'Remarques particulières concernant le financement 3
                'xlDestRange = xlWsheet2.Range("H230")
                'xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("SD" & nLigne)   'Y-a-t-il un autre financement (4) ?
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("A" & SautLigne).Value = xlWorkSheet.Range("SE" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("SE" & nLigne)   'Origine du financement 4 (préciser)
                'xlDestRange = xlWsheet2.Range("A231")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("C" & SautLigne).Value = xlWorkSheet.Range("SF" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("SF" & nLigne)   'Montant du financement 4 (en Ariary)
                'xlDestRange = xlWsheet2.Range("C231")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("SG" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("SG" & nLigne)   'Affectation du financement 4
                'xlDestRange = xlWsheet2.Range("D231")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("E" & SautLigne).Value = xlWorkSheet.Range("SH" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("SH" & nLigne)   'Date / périodicité du financement 4
                'xlDestRange = xlWsheet2.Range("E231")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("G" & SautLigne).Value = xlWorkSheet.Range("SI" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("SI" & nLigne)    'Administration / gestion des fonds 4 (qui gère ?)
                'xlDestRange = xlWsheet2.Range("G231")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("H" & SautLigne).Value = xlWorkSheet.Range("SJ" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("SJ" & nLigne)    'Remarques particulières concernant le financement 4
                'xlDestRange = xlWsheet2.Range("H231")
                'xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("SK" & nLigne)   'Y-a-t-il un autre financement (5) ?
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("A" & SautLigne).Value = xlWorkSheet.Range("SL" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("SL" & nLigne)   'Origine du financement 5 (préciser)
                'xlDestRange = xlWsheet2.Range("A232")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("C" & SautLigne).Value = xlWorkSheet.Range("SM" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("SM" & nLigne)    'Montant du financement 5 (en Ariary)
                'xlDestRange = xlWsheet2.Range("C232")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("SN" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("SN" & nLigne)    'Affectation du financement 5
                'xlDestRange = xlWsheet2.Range("D232")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("E" & SautLigne).Value = xlWorkSheet.Range("SO" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("SO" & nLigne)   'Date / périodicité du financement 5
                'xlDestRange = xlWsheet2.Range("E232")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("G" & SautLigne).Value = xlWorkSheet.Range("SP" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("SP" & nLigne)    'Administration / gestion des fonds 5 (qui gère ?)
                'xlDestRange = xlWsheet2.Range("G232")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("H" & SautLigne).Value = xlWorkSheet.Range("SQ" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("SQ" & nLigne)    'Remarques particulières concernant le financement 5
                'xlDestRange = xlWsheet2.Range("H232")
                'xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("SR" & nLigne)    'Y-a-t-il un autre financement (6) ?
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)
                ''xlWsheet2.Range("D15").Font.Size = 12

                SautLigne += 1
                xlWsheet2.Range("A" & SautLigne).Value = xlWorkSheet.Range("SS" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("SS" & nLigne)   'Origine du financement 6 (préciser)
                'xlDestRange = xlWsheet2.Range("A233")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("C" & SautLigne).Value = xlWorkSheet.Range("ST" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("ST" & nLigne)   'Montant du financement 6 (en Ariary)
                'xlDestRange = xlWsheet2.Range("C233")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("SU" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("SU" & nLigne)   'Affectation du financement 6
                'xlDestRange = xlWsheet2.Range("D233")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("E" & SautLigne).Value = xlWorkSheet.Range("SV" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("SV" & nLigne)   'Date / périodicité du financement 6
                'xlDestRange = xlWsheet2.Range("E233")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("G" & SautLigne).Value = xlWorkSheet.Range("SW" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("SW" & nLigne)   'Administration / gestion des fonds 6 (qui gère ?)
                'xlDestRange = xlWsheet2.Range("G233")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("H" & SautLigne).Value = xlWorkSheet.Range("SX" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("SX" & nLigne)    'Remarques particulières concernant le financement 6
                'xlDestRange = xlWsheet2.Range("H233")
                'xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("SY" & nLigne)    'Y-a-t-il un autre financement (7) ?
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("A" & SautLigne).Value = xlWorkSheet.Range("SZ" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("SZ" & nLigne)   'Origine du financement 7 (préciser)
                'xlDestRange = xlWsheet2.Range("A234")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("C" & SautLigne).Value = xlWorkSheet.Range("TA" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("TA" & nLigne)   'Montant du financement 7
                'xlDestRange = xlWsheet2.Range("C234")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("TB" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("TB" & nLigne)   'Affectation du financement 7
                'xlDestRange = xlWsheet2.Range("D234")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("E" & SautLigne).Value = xlWorkSheet.Range("TC" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("TC" & nLigne)    'Date / périodicité du financement 7
                'xlDestRange = xlWsheet2.Range("E234")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("G" & SautLigne).Value = xlWorkSheet.Range("TD" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("TD" & nLigne)    'Administration / gestion des fonds 7 (qui gère ?)
                'xlDestRange = xlWsheet2.Range("G234")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("H" & SautLigne).Value = xlWorkSheet.Range("TE" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("TE" & nLigne)    'Remarques particulières concernant le financement 7
                'xlDestRange = xlWsheet2.Range("H234")
                'xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("TF" & nLigne)   'Y-a-t-il un autre financement (8) ?
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("A" & SautLigne).Value = xlWorkSheet.Range("TG" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("TG" & nLigne)   'Origine du financement 8 (préciser)
                'xlDestRange = xlWsheet2.Range("A235")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("C" & SautLigne).Value = xlWorkSheet.Range("TH" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("TH" & nLigne)   'Montant du financement 8 (en Ariary)
                'xlDestRange = xlWsheet2.Range("C235")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("TI" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("TI" & nLigne)   'Affectation du financement 8
                'xlDestRange = xlWsheet2.Range("D235")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("E" & SautLigne).Value = xlWorkSheet.Range("TJ" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("TJ" & nLigne)   'Date / périodicité du financement 8
                'xlDestRange = xlWsheet2.Range("E235")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("G" & SautLigne).Value = xlWorkSheet.Range("TK" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("TK" & nLigne)   'Administration / gestion des fonds 8 (qui gère ?)
                'xlDestRange = xlWsheet2.Range("G235")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("H" & SautLigne).Value = xlWorkSheet.Range("TL" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("TL" & nLigne)   'Remarques particulières concernant le financement 8
                'xlDestRange = xlWsheet2.Range("H235")
                'xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("TM" & nLigne)   'Y-a-t-il un autre financement (9) ?
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("A" & SautLigne).Value = xlWorkSheet.Range("TN" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("TN" & nLigne)   'Origine du financement 9 (préciser)
                'xlDestRange = xlWsheet2.Range("A236")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("C" & SautLigne).Value = xlWorkSheet.Range("TO" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("TO" & nLigne)   'Montant du financement 9 (en Ariary)
                'xlDestRange = xlWsheet2.Range("C236")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("TP" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("TP" & nLigne)   'Affectation du financement 9
                'xlDestRange = xlWsheet2.Range("D236")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("E" & SautLigne).Value = xlWorkSheet.Range("TQ" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("TQ" & nLigne)    'Date / périodicité du financement 9
                'xlDestRange = xlWsheet2.Range("E236")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("G" & SautLigne).Value = xlWorkSheet.Range("TR" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("TR" & nLigne)   'Administration / gestion des fonds 9 (qui gère ?)
                'xlDestRange = xlWsheet2.Range("G236")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("H" & SautLigne).Value = xlWorkSheet.Range("TS" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("TS" & nLigne)   'Remarques particulières concernant le financement 9
                'xlDestRange = xlWsheet2.Range("H236")
                'xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("TT" & nLigne)   'Y-a-t-il un autre financement (10) ?
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                SautLigne += 1
                xlWsheet2.Range("A" & SautLigne).Value = xlWorkSheet.Range("TU" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("TU" & nLigne)   'Origine du financement 10 (préciser)
                'xlDestRange = xlWsheet2.Range("A237")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("C" & SautLigne).Value = xlWorkSheet.Range("TV" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("TV" & nLigne)   'Montant du financement 10 (en Ariary)
                'xlDestRange = xlWsheet2.Range("C237")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("D" & SautLigne).Value = xlWorkSheet.Range("TW" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("TW" & nLigne)   'Affectation du financement 10
                'xlDestRange = xlWsheet2.Range("D237")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("E" & SautLigne).Value = xlWorkSheet.Range("TX" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("TX" & nLigne)   'Date / périodicité du financement 10
                'xlDestRange = xlWsheet2.Range("E237")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("G" & SautLigne).Value = xlWorkSheet.Range("TY" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("TY" & nLigne)   'Administration / gestion des fonds 10 (qui gère ?)
                'xlDestRange = xlWsheet2.Range("G237")
                'xlSourceRange.Copy(xlDestRange)

                xlWsheet2.Range("H" & SautLigne).Value = xlWorkSheet.Range("TZ" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("TZ" & nLigne)   'Remarques particulières concernant le financement 10
                'xlDestRange = xlWsheet2.Range("H237")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne = 240
                xlWsheet2.Range("H" & SautLigne).Value = xlWorkSheet.Range("UA" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("UA" & nLigne)  'La gestion actuelle de l’infrastructure permet-elle de pérenniser la qualité du service fourni ?
                'xlDestRange = xlWsheet2.Range("H240")
                'xlSourceRange.Copy(xlDestRange)

                SautLigne += 2
                xlWsheet2.Range("H" & SautLigne).Value = xlWorkSheet.Range("UB" & nLigne).Text
                'xlSourceRange = xlWorkSheet.Range("UB" & nLigne)   'Remarque générale permettant de comprendre les raisons principales du bon fonctionnement / dysfonctionnement de l'infrastructure
                'xlDestRange = xlWsheet2.Range("H242")
                'xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("UC" & nLigne)  'Photo 1
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("UD" & nLigne)   '__version__
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("UE" & nLigne)   '_id
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("UF" & nLigne)   '_uuid
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("UG" & nLigne)    '_submission_time
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                ''xlSourceRange = xlWorkSheet.Range("UH" & nLigne)    '_index
                ''xlDestRange = xlWsheet2.Range("D15")
                ''xlSourceRange.Copy(xlDestRange)

                '***************************************************************************************

                'Mise en forme des rapports

                With xlWsheet2.Range("E1:G1")
                    .Merge()
                    .Font.Size = 18
                    .Font.Bold = True
                    .Font.Underline = False
                    .Font.Color = Color.Blue
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("E2:G2")
                    .Merge()
                    .Font.Size = 14
                    .Font.Bold = True
                    .Font.Underline = True
                    .Font.Color = Color.Blue
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With

                xlWsheet2.Range("B6:B18;F24:J30;G31:J33;E34:J38;H39:J41;H45:J45;E46:J48;D52:E62;I52:J62;D68:E73;I68:J69;J70:J72;H73:J73;H77:J77;C81:E82;B83:E85;D86:E87;B88:E94;J81;I82:J82;H83:J85;J86:J87;H88:J94;D98:E98;C99:E99;B100:E102;D103:E104;B105:E111;J98;I99:J99;H100:J102").Font.Size = 12
                'xlWsheet2.Range("B6:B18;F24:J30;G31:J33;E34:J38;H39:J41;H45:J45;E46:J48;D52:E62;I52:J62;D68:E73;I68:J69;J70:J72;H73:J73;H77:J77;C81:E82;B83:E85;D86:E87;B88:E94;J81;I82:J82;H83:J85;J86:J87;H88:J94;D98:E98;C99:E99;B100:E102;D103:E104;B105:E111;J98;I99:J99;H100:J102;
                '----------------J103:J104;H105:J111;C118:H125;H129;C132:H141;B147:G149;B152:F154;B157:E159;B162:D164;B169:I169;G173:I173;B176:C181;A189:J191;A197:J199;F206:I207;D212:I216;A221:J223;A228:J237;H240:J243").Font.Size = 12
                xlWsheet2.Range("J103:J104;H105:J111;C118:H125;H129;C132:H141;B147:G149;B152:F154;B157:E159;B162:D164;B169:I169;G173:I173;B176:C181;A189:J191;A197:J199;F206:I207;D212:I216;A221:J223;A228:J237;H240:J243").Font.Size = 12

                'tableau Mobilier

                With xlWsheet2.Range("C118:E125")
                    ' .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlEdgeLeft).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlEdgeTop).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlInsideHorizontal).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlInsideVertical).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With

                With xlWsheet2.Range("F118:G118")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("F119:G119")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("F120:G120")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("F121:G121")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("F122:G122")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("F123:G123")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("F124:G124")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("F125:G125")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With

                With xlWsheet2.Range("A122:B122")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlEdgeLeft).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("A123:B123")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlEdgeLeft).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("A124:B124")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlEdgeLeft).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("A125:B125")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlEdgeLeft).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With

                With xlWsheet2.Range("H118:H125")
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlInsideHorizontal).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With

                'tableau équipement didactique / pédagogique
                With xlWsheet2.Range("C132:E141")
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    '.Borders(Excel.XlBordersIndex.xlEdgeLeft).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlEdgeTop).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlInsideHorizontal).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlInsideVertical).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With

                With xlWsheet2.Range("H132:H141")
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlInsideHorizontal).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With

                With xlWsheet2.Range("F132:G132")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("F133:G133")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("F134:G134")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("F135:G135")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("F136:G136")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("F137:G137")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("F138:G138")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("F139:G139")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("F140:G140")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("F141:G141")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With

                With xlWsheet2.Range("A138:B138")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlEdgeLeft).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("A139:B139")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlEdgeLeft).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("A140:B140")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlEdgeLeft).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("A141:B141")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlEdgeLeft).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With

                'tableau Fréquentation de l'établissement
                With xlWsheet2.Range("B147:C147")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("B148:C148")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("B149:C149")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With

                With xlWsheet2.Range("D147:E147")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("D148:E148")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("D149:E149")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With

                With xlWsheet2.Range("F147:G147")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("F148:G148")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("F149:G149")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With

                With xlWsheet2.Range("B152:F154")
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlInsideHorizontal).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlInsideVertical).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With

                With xlWsheet2.Range("B157:E159")
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlInsideHorizontal).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlInsideVertical).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With

                With xlWsheet2.Range("B162:D164")
                    ' .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    '.Borders(Excel.XlBordersIndex.xlEdgeLeft).Weight = 2
                    '.Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = 2
                    '.Borders(Excel.XlBordersIndex.xlEdgeTop).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlInsideHorizontal).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlInsideVertical).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With

                'tableau personnel enseignant
                With xlWsheet2.Range("B169")
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With

                With xlWsheet2.Range("C169:D169")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With

                With xlWsheet2.Range("E169:G169")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    '.Borders(Excel.XlBordersIndex.xlEdgeLeft).Weight = 2
                    '.Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = 2
                    '.Borders(Excel.XlBordersIndex.xlEdgeTop).Weight = 2
                    '.Borders(Excel.XlBordersIndex.xlInsideHorizontal).Weight = 2
                    '.Borders(Excel.XlBordersIndex.xlInsideVertical).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With

                With xlWsheet2.Range("H169:I169")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With

                'tableau personnel non enseignant
                With xlWsheet2.Range("A176:A181")
                    '.Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlEdgeLeft).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = 2
                    '.Borders(Excel.XlBordersIndex.xlEdgeTop).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlInsideHorizontal).Weight = 2
                    '.Borders(Excel.XlBordersIndex.xlInsideVertical).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With

                With xlWsheet2.Range("B176:C176")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("B177:C177")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("B178:C178")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("B179:C179")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("B180:C180")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("B181:C181")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With

                'tableau maintenance effectuée
                With xlWsheet2.Range("A189:J191")
                    '.Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlEdgeLeft).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlEdgeTop).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlInsideHorizontal).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlInsideVertical).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("A189:B189")
                    .Merge()
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("A190:B190")
                    .Merge()
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("A191:B191")
                    .Merge()
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With

                With xlWsheet2.Range("C189:D189")
                    .Merge()
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("C190:D190")
                    .Merge()
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("C191:D191")
                    .Merge()
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With

                With xlWsheet2.Range("I189:J189")
                    .Merge()
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("I190:J190")
                    .Merge()
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("I191:J191")
                    .Merge()
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With

                'tableau maintenance prévue
                With xlWsheet2.Range("A197:J199")
                    '.Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlEdgeLeft).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlEdgeTop).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlInsideHorizontal).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlInsideVertical).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("A197:B197")
                    .Merge()
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("A198:B198")
                    .Merge()
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("A199:B199")
                    .Merge()
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With

                With xlWsheet2.Range("C197:D197")
                    .Merge()
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("C198:D198")
                    .Merge()
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("C199:D199")
                    .Merge()
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With

                With xlWsheet2.Range("I197:J197")
                    .Merge()
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("I198:J198")
                    .Merge()
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("I199:J199")
                    .Merge()
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With

                'tableau personnel
                With xlWsheet2.Range("A214:C214")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlEdgeLeft).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                End With
                With xlWsheet2.Range("A215:C215")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlEdgeLeft).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                End With
                With xlWsheet2.Range("A216:C216")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlEdgeLeft).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                End With

                With xlWsheet2.Range("D212:E212")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("D213:E213")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("D214:E214")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("D215:E215")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("D216:E216")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With

                With xlWsheet2.Range("F212:G212")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("F213:G213")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("F214:G214")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("F215:G215")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("F216:G216")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With

                With xlWsheet2.Range("H212:I212")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("H213:I213")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("H214:I214")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("H215:I215")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("H216:I216")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With

                'tableau réunions
                With xlWsheet2.Range("A221:J223")
                    '.Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlEdgeLeft).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = 2
                    '.Borders(Excel.XlBordersIndex.xlEdgeTop).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlInsideHorizontal).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlInsideVertical).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With

                With xlWsheet2.Range("D221:E221")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("D222:E222")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("D223:E223")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With

                With xlWsheet2.Range("G221:H221")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("G222:H222")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("G223:H223")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With

                With xlWsheet2.Range("I221:J221")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("I222:J222")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("I223:J223")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With

                'Tableau financement
                With xlWsheet2.Range("A228:J237")
                    '.Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlEdgeLeft).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlEdgeTop).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlInsideHorizontal).Weight = 2
                    .Borders(Excel.XlBordersIndex.xlInsideVertical).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With

                With xlWsheet2.Range("A228:B228")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("A229:B229")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("A230:B230")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("A231:B231")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("A232:B232")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("A233:B233")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("A234:B234")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("A235:B235")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("A236:B236")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("A237:B237")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With

                With xlWsheet2.Range("E228:F228")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("E229:F229")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("E230:F230")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("E231:F231")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("E232:F232")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("E233:F233")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("E234:F234")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("E235:F235")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("E236:F236")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("E237:F237")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With

                With xlWsheet2.Range("H228:J228")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("H229:J229")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("H230:J230")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("H231:J231")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("H232:J232")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("H233:J233")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("H234:J234")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("H235:J235")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("H236:J236")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                With xlWsheet2.Range("H237:J237")
                    .Merge()
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With
                ''------------------------------------
                'Dim eLigne As Integer = 0
                ''--------------------------------------
                '' - 6 - Tableau financement
                'For eLigne = 219 To 227
                '    If xlWsheet2.Range("D" & eLigne).Text.trim = "" Then
                '        xlWsheet2.Range("D" & eLigne).EntireRow.Delete()
                '    End If
                'Next eLigne

                '' - 5 - Tableau personnel enseignant
                'For eLigne = 167 To 171
                '    If xlWsheet2.Range("A" & eLigne).Text.trim = "" Then
                '        xlWsheet2.Range("A" & eLigne).EntireRow.Delete()
                '    End If
                'Next eLigne

                '' - 4 - Tableau Equipement didactique / pédagogique
                'For eLigne = 128 To 131
                '    If xlWsheet2.Range("C" & eLigne).Text.trim = "" Then
                '        xlWsheet2.Range("C" & eLigne).EntireRow.Delete()
                '    End If
                'Next eLigne

                '' - 3 - Tableau mobilier
                'For eLigne = 112 To 115
                '    If xlWsheet2.Range("C" & eLigne).Text.trim = "" Then
                '        xlWsheet2.Range("C" & eLigne).EntireRow.Delete()
                '    End If
                'Next eLigne

                '' - 2 - informations concernant les bâtiments 1 et 2
                'For eLigne = 78 To 84
                '    If xlWsheet2.Range("D" & eLigne).Text.trim = "" And xlWsheet2.Range("H" & eLigne).Text.trim = "" Then
                '        xlWsheet2.Range("D" & eLigne).EntireRow.Delete()
                '    End If
                'Next eLigne

                '' - 1 - localisation
                'For eLigne = 7 To 18
                '    If xlWsheet2.Range("B" & eLigne).Text.trim = "" Then
                '        xlWsheet2.Range("B" & eLigne).EntireRow.Delete()
                '    End If
                'Next eLigne
                'For eLigne = 8 To 13
                '    If xlWsheet2.Range("B" & eLigne).Text.trim = "" Then
                '        xlWsheet2.Range("A" & eLigne).Value = ""
                '    End If
                'Next eLigne

                MsgBox("Fin de parcours de la ligne : " & nLigne, MsgBoxStyle.Exclamation, "Vérification des lignes d'enquentes...")

                If Not File.Exists("C:\Program RIC\Adolphe Expert\RIC\Rapport Initial\" & nRapportEducation & "- District de " & nDistrictEducation & ".xlsx") Then

                    'nRapport = xlWorkSheet.Range("E" & nLigne).Text
                    'xlWorkBook2.SaveAs(Filename:="C:\Program RIC\Adolphe Expert\RIC\Rapport Initial\" & nRapport & "- District de " & nDistrict & ".xlsx", FileFormat:=51)
                    xlWorkBook2.SaveAs(Filename:="C:\Program RIC\Adolphe Expert\RIC\Rapport Initial\" & nRapportEducation & "- District de " & nDistrictEducation & ".xlsx")

                    'xlApp.Workbooks.Open("C:\Program RIC\Adolphe Expert\RIC\Rapport Initial\" & nRapport & "- District de " & nDistrict & ".xlsx")

                    '~~> Display Excel
                    xlApp.Visible = False
                Else
                    MsgBox("Le fichier " & nRapportEducation & "- District de " & nDistrictEducation & ".xlsx existe déjà !", MsgBoxStyle.Exclamation, "Oups ! Fichier en doublure...")
                    'nLigne = nLigne + 1
                    'MsgBox("Maintenant la ligne : " & nLigne, MsgBoxStyle.Exclamation, "Vérification des lignes d'enquentes...")
                End If

                'If xlWorkSheet.Range("UI" & nLigne).Text = "" Then
                'If xlWorkSheet.Range("A" & nLigne).Text = "" Then
                'If xlWorkSheet.Range("D" & 1 + nLigne).Text = "" Then

                'If xlWorkSheet.Range("D" & fLigne).Text = "" Then

                'If nRapportEducation = "" Then

                '    MsgBox("Fin de ligne trouvé", MsgBoxStyle.Exclamation, "Fin des lignes d'enquentes...")

                '    Exit For
                'End If

                MsgBox("On continue avec la ligne : " & nLigne + 1 & " ?", MsgBoxStyle.YesNo, "Toujours dans le boucle...")
                If MsgBoxResult.No = True Then
                    Exit For
                End If

                'If nRapportEducation = "" Then

            Else
                MsgBox("Fin de ligne trouvé", MsgBoxStyle.Exclamation, "Fin des lignes d'enquentes...")
                Exit For
            End If
            ' CLEAN UP. (CLOSE INSTANCES OF EXCEL OBJECTS.)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp) : xlApp = Nothing
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBook) : xlWorkBook = Nothing
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkSheet) : xlWorkSheet = Nothing

                ' CLEAN UP. (CLOSE INSTANCES OF EXCEL OBJECTS.)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBook2) : xlWorkBook2 = Nothing
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheet2) : xlWsheet2 = Nothing

                'xlWorkBook.Close() : xlApp.Quit()

                'Next nLigne

                'MsgBox("On a fini avec la ligne : " & nLigne, MsgBoxStyle.Exclamation, "Sorti de boucle...")

                'Catch ex As Exception
                'MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                'Finally
                '    '    'MAKE SURE TO KILL ALL INSTANCES BEFORE QUITING! if you fail to do this The service (excel.exe) will continue to run   
                NAR(xlWsheet2)
                'xlWorkBook2.Close(False)
                NAR(xlWorkBook)
            'xlApp.Workbooks.Close()
            'NAR(xlApp.Workbooks)
            'xlApp.Quit()
            NAR(xlApp)
            'VERY IMPORTANT   
            GC.Collect()
                'End Try

            Next nLigne

            MsgBox("On a fini avec la ligne : " & nLigne, MsgBoxStyle.Exclamation, "Sorti de boucle...")

    End Sub

    Private Sub ChargerRapportsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ChargerRapportsToolStripMenuItem.Click
        If Not File.Exists(sFileEducation) Then
            MsgBox("Veuillez charger les fichiers sources !", MsgBoxStyle.Exclamation, "Fichier source introuvable...")
            Exit Sub
        End If
        If Not File.Exists(mFileEducation) Then
            MsgBox("Veuillez charger les fichiers modèles !", MsgBoxStyle.Exclamation, "Fichier modèle introuvable...")
            Exit Sub
        Else MsgBox("Voullez-vous chargez les rapports initiaux maintenant ?", MsgBoxStyle.YesNo, "Chargement rapport initaial...")
            'If MsgBoxResult.Yes = True Then
            GenererRapportExcelInitialEducation(sFileEducation)
            'End If
        End If
    End Sub

    'End Class

    '////////////////////////////////////////////

End Class


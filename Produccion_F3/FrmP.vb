Imports System.Windows.Forms
Imports System.IO
Imports System.Data.SqlClient
Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Windows.Forms
Imports CrystalDecisions.ReportSource
Public Class FrmP
#Region "Variables"
    Public con As New Conexion
    Public itemcode As String
    Dim valBarCode As String
    Dim oCompany As SAPbobsCOM.Company
    Dim connectionString As String = Conexion.ObtenerConexion.ConnectionString
    Public Ready As Boolean
    Private Const CP_NOCLOSE_BUTTON As Integer = &H200
    Public Shared SQL_Conexion As SqlConnection = New SqlConnection()
    Public Shared ba As New List(Of String)
    Public Shared quantity As New List(Of Integer)
    Public Shared itemsimp As New List(Of String)
    Public Shared quantityimp As New List(Of Integer)
    Public Shared anchoimp As New List(Of Double)
    Public Shared wo As SAPbobsCOM.ProductionOrders
#End Region
    Protected Overloads Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim myCp As CreateParams = MyBase.CreateParams
            myCp.ClassStyle = myCp.ClassStyle Or CP_NOCLOSE_BUTTON
            Return myCp
        End Get
    End Property

    Public Sub New(ByVal user As String)
        MyBase.New()
        InitializeComponent()
        '  Note which form has called this one
        ToolStripStatusLabel1.Text = user
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        Dim SQL_da As SqlDataAdapter = New SqlDataAdapter("SELECT T0.DocNum FROM OWOR T0 where T0.Type = 'P' and T0.Status = 'P' and T0.DocNum LIKE '" + TextBox2.Text + "%' ORDER BY T0.DocNum", con.ObtenerConexion())
        Dim DT_dat As System.Data.DataTable = New System.Data.DataTable()
        SQL_da.Fill(DT_dat)
        DGV.DataSource = DT_dat
        con.ObtenerConexion.Close()
    End Sub

    Private Sub FrmP_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox2.Select()
        ComboBox1.DropDownStyle = ComboBoxStyle.DropDownList
        cargaORDER()
    End Sub

    Public Function cargaORDER()
        Dim SQL_da As SqlDataAdapter = New SqlDataAdapter("SELECT T0.DocNum FROM OWOR T0 where T0.Type = 'P' and T0.Status = 'P'", con.ObtenerConexion())
        Dim DT_dat As System.Data.DataTable = New System.Data.DataTable()
        SQL_da.Fill(DT_dat)
        DGV.DataSource = DT_dat
        con.ObtenerConexion.Close()
    End Function

    Private Sub DGV_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DGV.CellContentClick
        txtOrder.Text = DGV(0, DGV.CurrentCell.RowIndex).Value.ToString()
        Panel1.Visible = True
        DGV2.Visible = True
        Button4.Visible = True
        Button5.Visible = True
        DGV3.Visible = True
        Dim SQL_da As SqlDataAdapter = New SqlDataAdapter("SELECT T0.ItemCode, T2.ItemName, T0.PlannedQty FROM OWOR T0 inner join OITM T2 on T0.ItemCode = T2.ItemCode where T0.DocNum= '" + txtOrder.Text + "'", con.ObtenerConexion())
        Dim DT_dat As System.Data.DataTable = New System.Data.DataTable()
        SQL_da.Fill(DT_dat)
        Label3.Text = DT_dat.Rows(0).Item("ItemCode").ToString
        Label5.Text = DT_dat.Rows(0).Item("ItemName").ToString
        Label7.Text = DT_dat.Rows(0).Item("PlannedQty").ToString & " TM"
        con.ObtenerConexion.Close()

        Dim turnoAM_I As DateTime = CType("6:00:00 AM", DateTime)
        Dim turnoAM_F As DateTime = CType("6:00:00 PM", DateTime)

        Dim result As Integer = 0
        Dim result2 As Integer = 0
        result = DateTime.Compare(turnoAM_I, TimeOfDay.ToShortTimeString)
        result2 = DateTime.Compare(turnoAM_F, TimeOfDay.ToShortTimeString)
        If result = -1 Then
            If result2 = 1 Then
                ComboBox1.Text = "AM"
            Else
                ComboBox1.Text = "PM"
            End If
        End If
    End Sub

    Private Sub txtOrder_TextChanged(sender As Object, e As EventArgs) Handles txtOrder.TextChanged
        Dim SQL_da As SqlDataAdapter = New SqlDataAdapter("SELECT T0.ItemCode, T0.BaseQty, T0.U_ancho,T0.U_tiras,T0.U_peso,isnull(T0.LineNum,0) FROM WOR1 T0 where T0.[DocEntry] like '" + txtOrder.Text + "%'", con.ObtenerConexion())
        Dim DT_dat As System.Data.DataTable = New System.Data.DataTable()
        SQL_da.Fill(DT_dat)
        DGV2.DataSource = DT_dat
        con.ObtenerConexion.Close()
    End Sub

    Private Sub Update()
        Dim oReturn As Integer = -1
        Dim oError As Integer = 0
        Dim errMsg As String = ""
        Dim sql As String
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim objectCode As Integer

        Dim itemname As String
        Dim pesoreal As Double
        Dim postdate As String

        Try
            Dim result As Integer = MessageBox.Show("Desea Ingresar la Orden?", "Atencion", MessageBoxButtons.YesNoCancel)
            If result = DialogResult.Cancel Then
                MessageBox.Show("Cancelado")
            ElseIf result = DialogResult.No Then
                MessageBox.Show("No se realizara la orden")
            ElseIf result = DialogResult.Yes Then
                Dim cont As Integer = 0
                For Each row As DataGridViewRow In DGV3.Rows

                    If con.Connected = True Then
                        sql = ("select T0.DocEntry, T0.ItemCode, T1.itemName, T0.PlannedQty,T0.PostDate,T0.DueDate from OWOR T0 inner join OITM T1 on T0.itemcode = t1.itemcode where T0.Docnum = '" + DGV3.Rows(cont).Cells.Item(0).Value.ToString + "'")
                        oRecordSet = con.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRecordSet.DoQuery(sql)
                        If oRecordSet.RecordCount > 0 Then
                            objectCode = oRecordSet.Fields.Item(0).Value
                            itemcode = oRecordSet.Fields.Item(1).Value
                            itemname = oRecordSet.Fields.Item(2).Value
                            pesoreal = oRecordSet.Fields.Item(3).Value
                            postdate = oRecordSet.Fields.Item(4).Value
                        End If

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
                        oRecordSet = Nothing
                        GC.Collect()
                        'Dim items As String
                        'For Each row As DataGridViewRow In DGV2.Rows
                        '    items += row.Cells(0).Value.ToString + "|"
                        'Next
                        'imprime(items.TrimEnd("|"), itemname, "AnchoTira", pesoreal, items.TrimEnd("|"), "heat", "coil", docdate)




                        wo = con.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)
                        wo.GetByKey(objectCode)
                        wo.ProductionOrderStatus = SAPbobsCOM.BoProductionOrderStatusEnum.boposReleased
                        'wo.Update()
                        oReturn = wo.Update()
                        If oReturn <> 0 Then
                            MessageBox.Show(con.oCompany.GetLastErrorDescription)
                        Else

                            Dim items As String
                            sql = ("select T0.U_lotes, T0.U_tiras, T0.U_ancho from wor1 T0 where T0.DocEntry = '" + objectCode.ToString + "'")
                            'sql = ("SELECT DocEntry FROM WOR1 where DocEntry = " + objectCode + "")
                            oRecordSet = con.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet.DoQuery(sql)
                            Dim vueltas As Integer = Convert.ToInt32(oRecordSet.RecordCount)
                            Do While Not oRecordSet.EoF
                                itemsimp.Add(oRecordSet.Fields.Item(0).Value)
                                quantityimp.Add(oRecordSet.Fields.Item(1).Value)
                                anchoimp.Add(oRecordSet.Fields.Item(2).Value)
                                oRecordSet.MoveNext()
                            Loop

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
                            oRecordSet = Nothing
                            GC.Collect()

                            Dim conts As Integer

                            Do While vueltas > 0
                                conts = vueltas - 1
                                Dim cont1 As Integer
                                cont1 = 0
                                Do While cont1 < Convert.ToInt32(quantityimp.Item(conts))
                                    imprime(itemsimp.Item(conts), itemname, anchoimp.Item(conts), pesoreal, itemsimp.Item(conts), "heat", "coil", postdate, ComboBox1.Text)
                                    cont1 += cont1 + 1
                                Loop
                                vueltas = vueltas - 1
                            Loop


                            itemcode = String.Empty
                            items = String.Empty
                            Panel1.Visible = False
                            DGV2.Visible = True
                            cargaORDER()
                            DGV2.DataSource = Nothing
                            DGV3.Rows.Clear()
                            MessageBox.Show("Finalizado con Exito")
                        End If

                    End If
                    cont = cont + 1
                Next
                con.oCompany.Disconnect()
                con.oCompany = Nothing
            End If
        Catch ex As Exception
            MsgBox("Error: " + ex.Message.ToString)
            con.oCompany.Disconnect()
            con.oCompany = Nothing
        End Try
    End Sub
    Private Function FormatBarCode(code As String)
        Dim barcode As String = String.Empty
        barcode = String.Format("*{0}*", code)
        Return barcode
    End Function
    Private Sub imprime(bat As String, desc As String, anch As String, pes As String, bob As String, het As String, coi As String, ordr As String, turnos As String)
        Dim Report1 As New CrystalDecisions.CrystalReports.Engine.ReportDocument()
        Report1.PrintOptions.PaperOrientation = PaperOrientation.Portrait
        Report1.Load(Application.StartupPath + "\Report\Informe.rpt", CrystalDecisions.Shared.OpenReportMethod.OpenReportByDefault.OpenReportByDefault)
        Report1.SetParameterValue("CodBatch", itemcode)
        'Report1.SetParameterValue("CodBatch", txtBarcode.Text)
        Report1.SetParameterValue("descripcion", desc)
        Report1.SetParameterValue("anchotira", anch)
        Report1.SetParameterValue("pesoreal", pes)
        Report1.SetParameterValue("bobina", bob)
        Report1.SetParameterValue("heat", het)
        Report1.SetParameterValue("coil", coi)
        Report1.SetParameterValue("ordencorte", ordr)
        Report1.SetParameterValue("fechacorte", Now.ToShortDateString)
        Report1.SetParameterValue("turno", turnos)
        'CrystalReportViewer1.ReportSource = Report1
        Report1.PrintToPrinter(1, False, 0, 0)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            If con.Connected Then
                Update()
            Else
                con.MakeConnectionSAP()
                If con.Connected Then
                    Update()
                Else
                    MessageBox.Show("Error de Conexion, intente Nuevamente")
                End If
            End If
        Catch ex As Exception
            MsgBox("Error: " + ex.Message.ToString)
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim result As Integer = MessageBox.Show("Desea limpiar el objeto?", "Atencion", MessageBoxButtons.YesNoCancel)
        If result = DialogResult.Cancel Then
            MessageBox.Show("Cancelado")
        ElseIf result = DialogResult.No Then
            MessageBox.Show("Puede continuar!")
        ElseIf result = DialogResult.Yes Then
            TextBox2.Clear()
            txtOrder.Text = ""
            wo = Nothing
            DGV.DataSource = Nothing
            DGV2.DataSource = Nothing
            DGV3.Rows.Clear()
            MessageBox.Show("Inicie un objeto nuevo")
        End If
        Panel1.Visible = False
        DGV2.Visible = False
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        Dim existe As Boolean = DGV3.Rows.Cast(Of DataGridViewRow).Any(Function(x) CInt(x.Cells("Ordenes").Value) = DGV(0, DGV.CurrentCell.RowIndex).Value.ToString())
        If Not existe Then
            DGV3.Rows.Add(DGV(0, DGV.CurrentCell.RowIndex).Value.ToString())
        Else
            MessageBox.Show("Ya ingreso ese numero de orden")
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        DGV3.Rows.Remove(DGV3.CurrentRow)
    End Sub


End Class

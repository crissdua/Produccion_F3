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
    Public Shared itemcode As New List(Of String)
    Public Shared itemname As New List(Of String)
    Public Shared peso As New List(Of String)
    Public Shared comment As New List(Of String)
    Public Shared wo As SAPbobsCOM.ProductionOrders
#End Region
    Protected Overloads Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim myCp As CreateParams = MyBase.CreateParams
            myCp.ClassStyle = myCp.ClassStyle Or CP_NOCLOSE_BUTTON
            Return myCp
        End Get
    End Property

    Public Sub New()
        MyBase.New()
        InitializeComponent()
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
        Dim SQL_da As SqlDataAdapter = New SqlDataAdapter("SELECT T0.ItemCode, T2.ItemName, convert(varchar(10),convert(date,T0.DueDate,106),103) as DueDate , T0.U_Ancho_Tira, T0.PlannedQty FROM OWOR T0 inner join OITM T2 on T0.ItemCode = T2.ItemCode where T0.DocNum= '" + txtOrder.Text + "'", con.ObtenerConexion())
        Dim DT_dat As System.Data.DataTable = New System.Data.DataTable()
        SQL_da.Fill(DT_dat)
        Label3.Text = DT_dat.Rows(0).Item("ItemCode").ToString
        Label5.Text = DT_dat.Rows(0).Item("ItemName").ToString
        Label7.Text = DT_dat.Rows(0).Item("PlannedQty").ToString & " TM"
        Label10.Text = DT_dat.Rows(0).Item("DueDate").ToString
        Label12.Text = DT_dat.Rows(0).Item("U_Ancho_Tira").ToString
        con.ObtenerConexion.Close()

        Dim turnoAM_I As DateTime = CType("6:00:00 AM", DateTime)
        Dim turnoAM_F As DateTime = CType("6:00:00 PM", DateTime)

        Dim result As Integer = 0
        Dim result2 As Integer = 0
        result = DateTime.Compare(turnoAM_I, TimeOfDay.ToShortTimeString)
        result2 = DateTime.Compare(turnoAM_F, TimeOfDay.ToShortTimeString)
        If result = -1 Then
            If result2 = 1 Then
                ComboBox1.Text = "Turno 1"
            Else
                ComboBox1.Text = "Turno 2"
            End If
        End If
    End Sub

    Private Sub txtOrder_TextChanged(sender As Object, e As EventArgs) Handles txtOrder.TextChanged
        Dim SQL_da As SqlDataAdapter = New SqlDataAdapter("SELECT T0.ItemCode as 'Articulo',T0.PlannedQty as 'TM Requerido',T0.U_lotes as Lote, T0.U_ancho as Ancho,T0.U_tiras as 'Tiras a Reportar',T0.U_peso as 'Peso de Tira',isnull(T0.LineNum,0) as 'Linea' FROM WOR1 T0 where T0.[DocEntry] like '" + txtOrder.Text + "%'", con.ObtenerConexion())
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


        itemsimp.Clear()
        quantityimp.Clear()
        anchoimp.Clear()
        itemcode.Clear()
        itemname.Clear()
        peso.Clear()
        comment.Clear()

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
                        sql = ("select T0.DocEntry, T0.ItemCode, T0.PlannedQty,T0.PostDate,T0.DueDate from OWOR T0 inner join OITM T1 on T0.itemcode = t1.itemcode where T0.Docnum = '" + DGV3.Rows(cont).Cells.Item(0).Value.ToString + "'")
                        oRecordSet = con.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRecordSet.DoQuery(sql)
                        If oRecordSet.RecordCount > 0 Then
                            objectCode = oRecordSet.Fields.Item(0).Value
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
                            Dim itemcodeE As String
                            Dim itemnameE As String
                            Dim plannedqtyE As String
                            Dim duedateE As String
                            Dim anchotiraE As String
                            Dim SQL_da As SqlDataAdapter = New SqlDataAdapter("SELECT T0.ItemCode, T2.ItemName, T0.DueDate, T0.U_Ancho_Tira, T0.PlannedQty FROM OWOR T0 inner join OITM T2 on T0.ItemCode = T2.ItemCode where T0.DocNum= '" + txtOrder.Text + "'", con.ObtenerConexion())
                            Dim DT_dat As System.Data.DataTable = New System.Data.DataTable()
                            SQL_da.Fill(DT_dat)
                            itemcodeE = DT_dat.Rows(0).Item("ItemCode").ToString
                            itemnameE = DT_dat.Rows(0).Item("ItemName").ToString
                            plannedqtyE = DT_dat.Rows(0).Item("PlannedQty").ToString & " TM"
                            duedateE = DT_dat.Rows(0).Item("DueDate").ToString
                            anchotiraE = DT_dat.Rows(0).Item("U_Ancho_Tira").ToString
                            con.ObtenerConexion.Close()

                            'itemcodeE, itemnameE, plannedqtyE, duedateE, anchotiraE
                            Dim items As String
                            sql = ("select T0.U_lotes, T0.U_tiras, T0.U_ancho,T0.itemcode,T1.itemname,T0.U_peso,T0.U_comment from wor1 T0 inner join oitm T1 on T1.itemcode = t0.itemcode where T0.DocEntry = '" + objectCode.ToString + "'")
                            'sql = ("SELECT DocEntry FROM WOR1 where DocEntry = " + objectCode + "")
                            oRecordSet = con.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet.DoQuery(sql)
                            Dim vueltas As Integer = Convert.ToInt32(oRecordSet.RecordCount)
                            Do While Not oRecordSet.EoF
                                itemsimp.Add(oRecordSet.Fields.Item(0).Value)
                                quantityimp.Add(oRecordSet.Fields.Item(1).Value)
                                anchoimp.Add(oRecordSet.Fields.Item(2).Value)
                                itemcode.Add(oRecordSet.Fields.Item(3).Value)
                                itemname.Add(oRecordSet.Fields.Item(4).Value)
                                peso.Add(oRecordSet.Fields.Item(5).Value)
                                comment.Add(oRecordSet.Fields.Item(6).Value)
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
                                    If Convert.ToInt32(quantityimp.Item(conts)) >= 0 Then
                                        Dim SQL_heatcoil As SqlDataAdapter = New SqlDataAdapter("select  COALESCE(MAX(U_Coi), '0') as coil, COALESCE(MAX(U_Heat), '0') as heat from obtn where DistNumber ='" + itemsimp.Item(conts) + "'", con.ObtenerConexion())
                                        Dim DT_heatcoil As System.Data.DataTable = New System.Data.DataTable()
                                        SQL_heatcoil.Fill(DT_heatcoil)
                                        Dim heat As String
                                        Dim coil As String
                                        heat = DT_heatcoil.Rows(0).Item("coil").ToString
                                        coil = DT_heatcoil.Rows(0).Item("heat").ToString
                                        con.ObtenerConexion.Close()

                                        imprime(itemsimp.Item(conts) & "-" & cont1 + 1, itemname.Item(cont), anchoimp.Item(conts), peso.Item(cont), itemsimp.Item(conts), heat, coil, comment.Item(cont), ComboBox1.Text, itemcode.Item(cont), itemcodeE, itemnameE, plannedqtyE, duedateE, anchotiraE, txtOrder.Text)
                                        'itmcod As String, desc As String, anch As String, pes As String, batch As String, het As String, coi As String, comment As String, turnos As String
                                    End If
                                    cont1 = cont1 + 1
                                Loop
                                vueltas = vueltas - 1
                            Loop
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
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
            oRecordSet = Nothing
            GC.Collect()
        End Try
    End Sub
    Private Function FormatBarCode(code As String)
        Dim barcode As String = String.Empty
        barcode = String.Format("*{0}*", code)
        Return barcode
    End Function
    Private Sub imprime(itmcod As String, desc As String, anch As String, pes As String, batch As String, het As String, coi As String, comment As String, turnos As String, itmscod As String, itemcodeE As String, itemnameE As String, plannedqtyE As String, duedateE As String, anchotiraE As String, docnum As String)
        Dim Report1 As New CrystalDecisions.CrystalReports.Engine.ReportDocument()
        Report1.PrintOptions.PaperOrientation = PaperOrientation.Portrait
        Report1.Load(Application.StartupPath + "\Report\InformeF3.rpt", CrystalDecisions.Shared.OpenReportMethod.OpenReportByDefault.OpenReportByDefault)
        ''-----------------------------------------ENCABEZADO NO CAMBIA POR IMPRESION------------------------------------------
        Report1.SetParameterValue("itemcode", itemcodeE)
        Report1.SetParameterValue("u_ancho", anchotiraE)
        Report1.SetParameterValue("docnum", docnum)
        Report1.SetParameterValue("docdate", duedateE)
        Report1.SetParameterValue("descripcionEnc", itemnameE)
        Report1.SetParameterValue("pesoEnc", plannedqtyE)
        ''------------------------------------------DETALLE TRAE DATOS POR PARAMETROS------------------------------------------
        Report1.SetParameterValue("CodBatch", itmcod) 'col4
        Report1.SetParameterValue("descripcion", desc) 'col2
        Report1.SetParameterValue("pesoreal", pes) 'col5
        Report1.SetParameterValue("anchotira", anch) 'col6
        Report1.SetParameterValue("bobina", itmscod) 'col1
        Report1.SetParameterValue("heat", het)
        Report1.SetParameterValue("coil", coi)
        Report1.SetParameterValue("turno", turnos)
        Report1.SetParameterValue("comments", comment)
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

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim result As Integer = MessageBox.Show("Desea salir del modulo?", "Atencion", MessageBoxButtons.YesNo)
        If result = DialogResult.No Then
            MessageBox.Show("Puede continuar")
        ElseIf result = DialogResult.Yes Then
            Try
                con.oCompany.Disconnect()
            Catch
            End Try
            Me.Hide()
        End If
    End Sub

    Private Sub Panel1_Paint(sender As Object, e As PaintEventArgs) Handles Panel1.Paint

    End Sub
End Class

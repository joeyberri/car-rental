Imports System.Drawing.Printing
Imports System.Data.DataTable
Imports Excel = Microsoft.office.Interop.Excel
Imports Microsoft.Office
Imports Microsoft.Office.Interop
Imports System.IO
Public Class Form1

    Private Sub TabPage1_Click(sender As Object, e As EventArgs) Handles TabPage1.Click

    End Sub

    Private Sub BtnAddVehicle_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles rbtInteriorYes.CheckedChanged

        If rbtInteriorYes.Checked Then
            dgvInteriorExtra.Enabled = True
            btnAddInteriorReceipt.Enabled = True
        End If
    End Sub

    Private Sub SaveToExcel()

        Dim excel As Microsoft.Office.Interop.Excel._Application = New Microsoft.Office.Interop.Excel.Application()
        Dim workbook As Microsoft.Office.Interop.Excel._Workbook = excel.Workbooks.Add(Type.Missing)
        Dim worksheet As Microsoft.Office.Interop.Excel._Worksheet = Nothing

        Try

            worksheet = workbook.ActiveSheet
            worksheet.Name = "ExportedFromDatGrid"

            Dim cellrowindex As Integer = 1
            Dim cellColumnindex As Integer = 1

            For j As Integer = 0 To dgvReceipt.Columns.Count - 1
                worksheet.Cells(cellrowindex, cellColumnindex) = dgvReceipt.Columns(j).HeaderText
                cellColumnindex += 1
            Next

            cellColumnindex = 1
            cellrowindex += 1
            For i As Integer = 0 To dgvReceipt.Rows.Count - 2
                For j As Integer = 0 To dgvReceipt.Columns.Count - 1
                    worksheet.Cells(cellrowindex, cellColumnindex) = dgvReceipt.Rows(i).Cells(j).Value.ToString()
                    cellColumnindex += 1
                Next
                cellColumnindex = 1
                cellrowindex += 1
            Next

            Dim saveDialog As New SaveFileDialog()
            saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
            saveDialog.FilterIndex = 2

            If saveDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                workbook.SaveAs(saveDialog.FileName)
                MessageBox.Show("Exported Successful")
            End If

        Catch ex As System.Exception
            MessageBox.Show(ex.Message)
        Finally
            excel.Quit()
            workbook = Nothing
            excel = Nothing
        End Try
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles rbtInteriorNone.CheckedChanged
        If rbtInteriorNone.Checked Then
            dgvInteriorExtra.Enabled = False
            btnAddInteriorReceipt.Enabled = False
        End If
    End Sub

    Private Sub RbtExteriorNone_CheckedChanged(sender As Object, e As EventArgs) Handles rbtExteriorNone.CheckedChanged
        If rbtExteriorNone.Checked Then
            dgvExteriorExtra.Enabled = False
            btnAddExteriorReceipt.Enabled = False
        End If
    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles rbtExteriorYes.CheckedChanged
        If rbtExteriorYes.Checked Then
            dgvExteriorExtra.Enabled = True
            btnAddExteriorReceipt.Enabled = True
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim iExit As DialogResult
        iExit = MessageBox.Show("Do you really want to exit?", "Close", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)
        If iExit = DialogResult.Yes Then
            Application.Exit()
        End If
    End Sub

    Private Sub BtnAddVehicle_Click_1(sender As Object, e As EventArgs) Handles btnAddVehicle.Click
        dgvVehicle.Rows.Add(txtLotNumber.Text, txtMake.Text, txtModel.Text, nudYear.Text, txtMileage.Text, txtEngineCapacity.Text,
                        nudPrice.Text)
        MessageBox.Show("Vehicle added successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

    End Sub

    Private Sub BtnAddExtra_Click(sender As Object, e As EventArgs) Handles btnAddExtra.Click

        If cmbExtraType.SelectedIndex = 0 Then
            dgvInteriorExtra.Rows.Add(txtExtraID.Text, txtExtraName.Text, nudExtraCost.Text)
            MessageBox.Show("Extra added successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        ElseIf cmbExtraType.SelectedIndex = 1 Then
            dgvExteriorExtra.Rows.Add(txtExtraID.Text, txtExtraName.Text, nudExtraCost.Text)
            MessageBox.Show("Extra added successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

    End Sub

    Private Sub SaveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MenuStripPrint.Click
        IPrint()
    End Sub

    Private Sub ExitToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles MenuStripExit.Click
        Dim iExit As DialogResult
        iExit = MessageBox.Show("Do you really want to exit?", "Close", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)
        If iExit = DialogResult.Yes Then
            Application.Exit()
        End If
    End Sub

    Private Sub FileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FileToolStripMenuItem.Click

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles btnAddVehicleReceipt.Click
        If dgvVehicle.CurrentRow Is Nothing Or dgvVehicle.CurrentCell.Value = "" Then
            MessageBox.Show("You have to select a vehicle first", "Sorry", MessageBoxButtons.OKCancel, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        Else

            dgvReceipt.Rows.Add(dgvVehicle.CurrentRow.Cells(0).Value.ToString, "Vehicle", dgvVehicle.CurrentRow.Cells(1).Value.ToString & " " & dgvVehicle.CurrentRow.Cells(2).Value.ToString, dgvVehicle.CurrentRow.Cells(6).Value.ToString)
        End If
    End Sub

    Private Sub BtnAddInteriorReceipt_Click(sender As Object, e As EventArgs) Handles btnAddInteriorReceipt.Click
        If dgvReceipt.RowCount = 0 Then
            MessageBox.Show("You have to select a vehicle first", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        ElseIf dgvExteriorExtra.CurrentRow Is Nothing Then
            MessageBox.Show("You have to select an exterior extra first", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        Else
            dgvReceipt.Rows.Add(dgvInteriorExtra.CurrentRow.Cells(0).Value.ToString, "Interior Extra", dgvInteriorExtra.CurrentRow.Cells(1).Value.ToString, dgvInteriorExtra.CurrentRow.Cells(2).Value.ToString)
        End If
    End Sub

    Private Sub BtnAddExteriorReceipt_Click(sender As Object, e As EventArgs) Handles btnAddExteriorReceipt.Click
        If dgvReceipt.RowCount = 0 Then
            MessageBox.Show("You have to select a vehicle first", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)

        ElseIf dgvExteriorExtra.CurrentRow Is Nothing Then
            MessageBox.Show("You have to select an exterior extra first", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        Else
            dgvReceipt.Rows.Add(dgvExteriorExtra.CurrentRow.Cells(0).Value.ToString, "Exterior Extra", dgvExteriorExtra.CurrentRow.Cells(1).Value.ToString, dgvExteriorExtra.CurrentRow.Cells(2).Value.ToString)
        End If
    End Sub

    Private bitmap As Bitmap

    Private Sub IPrint()
        Dim height As Integer = dgvReceipt.Height
        dgvReceipt.Height = dgvReceipt.RowCount * dgvReceipt.RowTemplate.Height
        bitmap = New Bitmap(Me.dgvReceipt.Width, Me.dgvReceipt.Height)
        dgvReceipt.DrawToBitmap(bitmap, New Rectangle(0, 0, Me.dgvReceipt.Width, Me.dgvReceipt.Height))
        PrintPreviewDialog1.Document = PrintDocument1
        PrintPreviewDialog1.PrintPreviewControl.Zoom = 1
        PrintPreviewDialog1.ShowDialog()
        dgvReceipt.Height = height
    End Sub

    Private Sub BtnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        IPrint()
    End Sub


    Private Sub PrintDocument1_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        e.Graphics.DrawImage(bitmap, 0, 0)
    End Sub

    ''

    Private arrayProductID() As String = {"ABCD001", "ABCD002", "ABCD003", "ABCD004", "ABCD005", "ABCD006", "ABCD007"}
    Private arrayProductName() As String = {"Engine oil", "Fan Belt", "Oil Filter", "Wiper Blades", "Brake Fluids", "Coolant", "Washer Fluids"}
    Private arrayPrice() As Double = {150, 220, 430, 90, 56, 390, 22}
    Private index As Integer
    Dim newIndex As Integer = 1
    Dim totalAmount, discountAmount, taxableAmount, taxAmount, payableAmount As Double

    Private Sub BtnRemoveItem_Click(sender As Object, e As EventArgs) Handles btnRemoveItem.Click
        Dim userchoice As Double
        If dgvReceipt.Rows.Count > 1 Then
            index = dgvReceipt.CurrentRow.Index
            userchoice = dgvReceipt.Rows(index).Cells(5).Value
            dgvReceipt.Rows.Remove(dgvReceipt.CurrentRow)
            totalAmount = totalAmount - userchoice

            discountAmount = CalculateDiscount(totalAmount)
            taxableAmount = totalAmount - discountAmount
            taxAmount = CalculateTax(taxableAmount)
            payableAmount = taxableAmount + taxAmount
            DisplayResult()
        End If
    End Sub

    Private Sub TxtVehicleCount_TextChanged(sender As Object, e As EventArgs) Handles txtVehicleCount.TextChanged
        Dim i As Integer = 0
        Dim j As Integer = 0
        While j < dgvReceipt.RowCount - 1
            If dgvReceipt.Rows(index).Cells(1).Value = "Vehicle" Then i += 1
            txtVehicleCount.Text = i.ToString

        End While


    End Sub




    Private Function CalculateDiscount(ByVal receiptTotal As Double) As Double
        Dim discount As Double
        If receiptTotal <= 200 Then
            discount = 0.02 * receiptTotal
        ElseIf receiptTotal <= 1000 Then
            discount = 0.05 * receiptTotal
        ElseIf receiptTotal > 1000 Then
            discount = 0.1 * receiptTotal
        End If
        Return discount
    End Function

    Private Function CalculateTax(ByVal taxable As Double) As Double
        Dim taxAmount As Double
        taxAmount = 0.12 * taxable
        Return taxAmount
    End Function

    Private Sub DisplayResult()
        txtTotal.Text = totalAmount.ToString("C2")
        txtDiscount.Text = discountAmount.ToString("C2")
        txtSalesTax.Text = taxableAmount.ToString("C2")
        txtSubtotal.Text = payableAmount.ToString("C2")
    End Sub


    Private Sub ClearControls()
        totalAmount = 0.0
        newIndex = 1
        dgvReceipt.Rows.Clear()
        txtTotal.Text = ""
        txtSalesTax.Text = ""
        txtDiscount.Text = ""
        txtGrandTotalDue.Text = ""
        txtSubtotal.Text = ""
        txtVehicleCount.Text = ""
    End Sub
End Class

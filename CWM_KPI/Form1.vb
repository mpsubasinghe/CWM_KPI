Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Data.OleDb

Public Class Form1
    Dim Xds As DataSet
    Dim ds As DataSet
    Dim ds1 As DataSet
    Dim ds2 As DataSet
    Dim ds3 As DataSet
    Dim ds4 As DataSet

    Dim ds5 As DataSet
    Dim ds6 As DataSet
    Dim ds7 As DataSet
    Dim ds8 As DataSet

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click


        Dim oXL As Excel.Application
        Dim oWB As Excel.Workbook
        Dim oSheet As Excel.Worksheet
        Dim chartRange As Excel.Range
        oXL = CreateObject("Excel.Application")
        oWB = oXL.Workbooks.Add
        oSheet = oWB.ActiveSheet
        Dim s As Integer = 0

        With oWB
            .Sheets("Sheet1").Select()
            .Sheets(1).Name = "SalesReport"
        End With
        s = s + 2
        oSheet.Cells(s, 1).value = "C.W.Mackie PLC."
        chartRange = oSheet.Range("A" & s, "H" & s)
        chartRange.Merge()
        chartRange.HorizontalAlignment = 3
        chartRange.VerticalAlignment = 3
        chartRange.Font.Bold = True
        chartRange.Font.Size = 14

        s = s + 1
        oSheet.Cells(s, 1).value = "SFA KPI Document as at  - " & Format(Dat.Value.Date, "dd/MM/yyyy")
        chartRange = oSheet.Range("A" & s, "H" & s)
        chartRange.Merge()
        chartRange.HorizontalAlignment = 3
        chartRange.VerticalAlignment = 3
        chartRange.Font.Bold = True
        chartRange.Font.Italic = True
        chartRange.Font.Size = 12
        's = s + 1
        'oSheet.Cells(s, 1).value = "Free Issue Summary ( Division : " & ComboBoxDiv.Text & " ) "
        'oSheet.Cells(s, 1).Font.Bold = True

        's = s + 1
        'oSheet.Cells(s, 1).value = "Period From " & DateTimePicker1.Value.Date.ToString("yyyy-MM-dd") & " To " & DateTimePicker2.Value.Date.ToString("yyyy-MM-dd")
        'oSheet.Cells(s, 1).Font.Bold = True


        s = s + 2
        oSheet.Cells(s, 1).value = "Manager / Territory"

        oSheet.Cells(s, 2).value = "Day Sale (Below 50,000)"
        oSheet.Cells(s, 3).value = "PPC (Below 40)"
        oSheet.Cells(s, 4).value = "First Invoice After 9:30 AM"
        oSheet.Cells(s, 5).value = "Second Invoice After 9:30 AM"
        oSheet.Cells(s, 6).value = "Last Invoice Before 3:30 PM"
        oSheet.Cells(s, 7).value = "Returns Over 5000 (Sound)"
        oSheet.Cells(s, 8).value = "No of Sound Returns"
        oSheet.Rows("5:5").rowheight = 50
        oSheet.Columns("A:V").AutoFit()
        oSheet.Columns("A:V").WrapText = True

        chartRange = oSheet.Range(oSheet.Cells(s, 1), oSheet.Cells(s, 8))
        chartRange.Font.Bold = True

        oSheet.Columns("D:D").NumberFormat = "@"
        oSheet.Columns("E:E").NumberFormat = "@"
        oSheet.Columns("F:F").NumberFormat = "@"

        ds6 = GetDataACC("", "p2", "p3", "p4", "p5", "@dt", "pn2", "pn3", "pn4", "pn5", "SELECT Area.AreaNme, Area.AreaCode FROM  Area   GROUP BY Area.AreaNme, Area.AreaCode ORDER BY Area.AreaNme")
        'MsgBox("neww load with new rep" & ds1.Tables(0).Rows.Count)
        If ds6.Tables(0).Rows.Count > 0 Then
            ProgressBar1.Value = 0
            ProgressBar1.Maximum = ds6.Tables(0).Rows.Count
            For a As Integer = 0 To ds6.Tables(0).Rows.Count - 1
                ProgressBar1.Value = ProgressBar1.Value + 1
                s = s + 2
                oSheet.Cells(s, 1).value = ds6.Tables(0).Rows(a)("AreaNme").ToString()
                ' oSheet.Cells(s, 1).ForeColor = Color.Blue
                'oSheet.Cells(s, 1).Font.Size = 10
                oSheet.Cells(s, 1).Font.Bold = True
                ds2 = GetDataACC("", "p2", "p3", "p4", "p5", "@dt", "pn2", "pn3", "pn4", "pn5", "SELECT  Sector.Sector, Sector.Sect_Code FROM Sector WHERE  AraCode='" & ds6.Tables(0).Rows(a)("AreaCode").ToString() & "' GROUP BY Sector.Sector, Sector.Sect_Code ORDER BY Sector.Sector")
                'MsgBox("neww load with new rep" & ds1.Tables(0).Rows.Count)
                If ds2.Tables(0).Rows.Count > 0 Then
                    ProgressBar2.Value = 0
                    ProgressBar2.Maximum = ds2.Tables(0).Rows.Count
                    For b As Integer = 0 To ds2.Tables(0).Rows.Count - 1
                        ProgressBar2.Value = ProgressBar2.Value + 1
                        s = s + 1
                        oSheet.Cells(s, 1).value = ds2.Tables(0).Rows(b)("Sector").ToString()

                        ds3 = GetDataACC("", "p2", "p3", "p4", "p5", "@dt", "pn2", "pn3", "pn4", "pn5", "SELECT Sum([TrAmt]-[TrDisa]+[TrGST]) AS Expr1, Drill.DatInv From Drill Where (((Drill.Active) = True) And ((Drill.Sector) = " & ds2.Tables(0).Rows(b)("Sect_Code").ToString() & "))  GROUP BY Drill.DatInv HAVING (((Sum([TrAmt]-[TrDisa]+[TrGST]))<50000) AND ((Drill.DatInv)=" & Format(Dat.Value, "yyyyMMdd") & "))")
                        'MsgBox("neww load with new rep" & ds1.Tables(0).Rows.Count)
                        If ds3.Tables(0).Rows.Count > 0 Then
                            oSheet.Cells(s, 2).value = ds3.Tables(0).Rows(ds3.Tables(0).Rows.Count - 1)("Expr1").ToString()
                            oSheet.Cells(s, 2).NumberFormat = "#,###,###"
                        Else
                            oSheet.Cells(s, 2).value = ""
                        End If

                        ds4 = GetDataACC("", "p2", "p3", "p4", "p5", "@dt", "pn2", "pn3", "pn4", "pn5", "SELECT Last(Drill.InvNo) AS LastOfInvNo, Sum([TrAmt]+[TrGST]-[TrDisa]) AS Exp, Drill.Retailer From Drill WHERE (((Drill.ShortExp)<>'Y')) GROUP BY Drill.DatInv, Drill.Sector, Drill.Retailer HAVING (((Drill.DatInv)=" & Format(Dat.Value, "yyyyMMdd") & ") AND ((Sum([TrAmt]+[TrGST]-[TrDisa]))>0) AND ((Drill.Sector)=" & ds2.Tables(0).Rows(b)("Sect_Code").ToString() & ") )")
                        ds5 = GetDataACC("", "p2", "p3", "p4", "p5", "@dt", "pn2", "pn3", "pn4", "pn5", "SELECT Last(Drill.InvNo) AS LastOfInvNo, Sum([TrAmt]+[TrGST]-[TrDisa]) AS Exp, Drill.Retailer From Drill WHERE (((Drill.ShortExp)<>'Y')) GROUP BY Drill.DatInv, Drill.Sector, Drill.Retailer HAVING (((Drill.DatInv)=" & Format(Dat.Value, "yyyyMMdd") & ") AND ((Sum([TrAmt]+[TrGST]-[TrDisa]))<=0) AND ((Drill.Sector)=" & ds2.Tables(0).Rows(b)("Sect_Code").ToString() & ") )")
                        Dim PPCtem As Integer = 0

                        PPCtem = ds4.Tables(0).Rows.Count - ds5.Tables(0).Rows.Count

                        If PPCtem < 40 Then
                            oSheet.Cells(s, 3).value = PPCtem
                        Else
                            oSheet.Cells(s, 3).value = ""
                        End If

                        oSheet.Cells(s, 4).value = getTime("First", " 09:30:00 AM", ds2.Tables(0).Rows(b)("Sect_Code").ToString())
                        oSheet.Cells(s, 5).value = getTime("Second", " 09:30:00 AM", ds2.Tables(0).Rows(b)("Sect_Code").ToString())
                        oSheet.Cells(s, 6).value = getTime("Last", " 03:30:00 PM", ds2.Tables(0).Rows(b)("Sect_Code").ToString())

                        ds3 = GetDataACC("", "p2", "p3", "p4", "p5", "@dt", "pn2", "pn3", "pn4", "pn5", "SELECT Sum([TrAmt]-[TrDisa]+[TrGST]) AS Expr1, Drill.PrtID From Drill  WHERE (((Drill.ShortExp)<>'Y')) GROUP BY Drill.DatInv, Drill.Sector, Drill.PrtID HAVING (((Drill.DatInv)=" & Format(Dat.Value, "yyyyMMdd") & ") AND ((Drill.Sector)=" & ds2.Tables(0).Rows(b)("Sect_Code").ToString() & ") AND ((Drill.PrtID)='S'))")
                        'MsgBox("neww load with new rep" & ds1.Tables(0).Rows.Count)
                        If ds3.Tables(0).Rows.Count > 0 Then
                            If (ds3.Tables(0).Rows(ds3.Tables(0).Rows.Count - 1)("Expr1") * -1) >= 5000 Then
                                oSheet.Cells(s, 7).value = (ds3.Tables(0).Rows(ds3.Tables(0).Rows.Count - 1)("Expr1") * -1)
                                oSheet.Cells(s, 7).NumberFormat = "#,###,###"
                            Else
                                oSheet.Cells(s, 7).value = ""
                            End If
                        End If



                        ds3 = GetDataACC("", "p2", "p3", "p4", "p5", "@dt", "pn2", "pn3", "pn4", "pn5", "SELECT Drill.InvNo From Drill  WHERE (((Drill.ShortExp)<>'Y')) GROUP BY Drill.Sector, Drill.DatInv, Drill.PrtID, Drill.InvNo HAVING (((Drill.Sector)=" & ds2.Tables(0).Rows(b)("Sect_Code").ToString() & ") AND ((Drill.DatInv)=" & Format(Dat.Value, "yyyyMMdd") & ") AND ((Drill.PrtID)='S'))")
                        'MsgBox("neww load with new rep" & ds1.Tables(0).Rows.Count)
                        If ds3.Tables(0).Rows.Count > 0 Then
                            oSheet.Cells(s, 8).value = ds3.Tables(0).Rows.Count
                        End If

                    Next
                End If
            Next
        End If





        oSheet.Columns("A:A").ColumnWidth = 20
        oSheet.Columns("B:B").ColumnWidth = 10
        oSheet.Columns("C:C").ColumnWidth = 8
        oSheet.Columns("D:D").ColumnWidth = 20
        oSheet.Columns("E:E").ColumnWidth = 20
        oSheet.Columns("F:F").ColumnWidth = 20
        oSheet.Columns("G:G").ColumnWidth = 14
        oSheet.Columns("H:H").ColumnWidth = 8

        oSheet.Range(oSheet.Cells(5, 1), oSheet.Cells(s, 8)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
        oSheet.Range(oSheet.Cells(5, 1), oSheet.Cells(s, 8)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
        oSheet.Range(oSheet.Cells(5, 1), oSheet.Cells(s, 8)).Borders(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlContinuous
        oSheet.Range(oSheet.Cells(5, 1), oSheet.Cells(s, 8)).Borders(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlContinuous
        oSheet.Range(oSheet.Cells(5, 1), oSheet.Cells(s, 8)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
        oSheet.Range(oSheet.Cells(5, 1), oSheet.Cells(s, 8)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous


        'oSheet.Columns.AutoFit()

        oXL.Visible = True
        oXL.UserControl = True

        ' Make sure that you release object references.

        oSheet = Nothing
        oWB = Nothing
        ' oXL.Quit()
        oXL = Nothing
    End Sub
    Function getTime(status As String, stime As String, sector As String) As String
        Dim d1 As String = Dat.Value.Date.ToString("MM/dd/yyyy") & stime

        If status = "First" Then
            ds1 = GetDataACC("", "p2", "p3", "p4", "p5", "@dt", "pn2", "pn3", "pn4", "pn5", "SELECT Drill.InvNo, Drill.DatTme From Drill GROUP BY Drill.InvNo, Drill.DatInv, Drill.Sector, Drill.DatTme Having (((Drill.DatInv) = " & Format(Dat.Value, "yyyyMMdd") & ") And ((Drill.Sector) = " & sector & ")) ORDER BY Drill.InvNo")
            'MsgBox("neww load with new rep" & ds1.Tables(0).Rows.Count)
            If ds1.Tables(0).Rows.Count > 0 Then

                If Convert.ToDateTime(ds1.Tables(0).Rows(0)("DatTme")) > d1 Then
                    Return ds1.Tables(0).Rows(0)("DatTme")
                Else
                    Return ""
                End If

            End If
        End If

        If status = "Second" Then
            ds1 = GetDataACC("", "p2", "p3", "p4", "p5", "@dt", "pn2", "pn3", "pn4", "pn5", "SELECT Drill.InvNo, Drill.DatTme From Drill GROUP BY Drill.InvNo, Drill.DatInv, Drill.Sector, Drill.DatTme Having (((Drill.DatInv) = " & Format(Dat.Value, "yyyyMMdd") & ") And ((Drill.Sector) = " & sector & ")) ORDER BY Drill.InvNo")
            'MsgBox("neww load with new rep" & ds1.Tables(0).Rows.Count)
            If ds1.Tables(0).Rows.Count > 0 Then

                If Convert.ToDateTime(ds1.Tables(0).Rows(1)("DatTme")) > d1 Then
                    Return ds1.Tables(0).Rows(1)("DatTme")
                Else
                    Return ""
                End If

            End If
        End If

        If status = "Last" Then
            ds1 = GetDataACC("", "p2", "p3", "p4", "p5", "@dt", "pn2", "pn3", "pn4", "pn5", "SELECT Drill.InvNo, Drill.DatTme From Drill GROUP BY Drill.InvNo, Drill.DatInv, Drill.Sector, Drill.DatTme Having (((Drill.DatInv) = " & Format(Dat.Value, "yyyyMMdd") & ") And ((Drill.Sector) = " & sector & ")) ORDER BY Drill.InvNo")
            'MsgBox("neww load with new rep" & ds1.Tables(0).Rows.Count)
            If ds1.Tables(0).Rows.Count > 0 Then

                If Convert.ToDateTime(ds1.Tables(0).Rows(ds1.Tables(0).Rows.Count - 1)("DatTme")) < d1 Then
                    Return ds1.Tables(0).Rows(ds1.Tables(0).Rows.Count - 1)("DatTme")
                Else
                    Return ""
                End If

            End If
        End If

    End Function
End Class

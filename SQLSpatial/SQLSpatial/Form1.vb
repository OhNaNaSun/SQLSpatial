Imports System.Data.SqlClient
Imports Microsoft.SqlServer.Types
Public Class Form1

    Private pts, plines, regions As New ArrayList(0)
    Private xmin, xmax, ymin, ymax As Double

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim server As String = TextBox1.Text
        Dim database As String = TextBox2.Text
        Dim table As String = TextBox3.Text
        Dim builder As New SqlConnectionStringBuilder() '引用类型
        builder.DataSource = server
        builder.InitialCatalog = database
        builder.IntegratedSecurity = True
        'Dim connection As New SqlConnection(builder.ConnectionString)
        Dim command As String = "select * from " & table
        Dim adapter As New SqlDataAdapter(command, builder.ConnectionString)

        Try
            'connection.Open()
            'MsgBox("连接成功")
            'connection.Close()
            If DataSet1.Tables.Contains(table) = False Then
                DataSet1.Tables.Add(table)
            Else
                DataSet1.Tables(table).Clear()
            End If
            adapter.Fill(DataSet1, table)
            DataGridView1.DataSource = DataSet1.Tables(table)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        read()
    End Sub

    Private Sub read()
        xmin = Double.PositiveInfinity
        ymin = Double.PositiveInfinity
        xmax = Double.NegativeInfinity
        ymax = Double.NegativeInfinity

        pts.Clear()
        plines.Clear()
        regions.Clear()

        For i As Integer = 0 To DataGridView1.RowCount - 2
            Dim geom As SqlGeometry = CType(DataGridView1.Rows(i).Cells("geom").Value, SqlGeometry)
            Select Case geom.STGeometryType.Value
                Case "Point"
                    Dim pt As PointF
                    pt.X = geom.STX.Value
                    pt.Y = geom.STY.Value
                    xmin = Math.Min(xmin, pt.X)
                    ymin = Math.Min(ymin, pt.Y)
                    xmax = Math.Max(xmax, pt.X)
                    ymax = Math.Max(ymax, pt.Y)
                    pts.Add(pt)
                Case "LineString"
                    Dim pline As New ArrayList(1)
                    Dim points(geom.STNumPoints.Value - 1) As PointF
                    For j As Integer = 0 To geom.STNumPoints.Value - 1
                        points(j).X = geom.STPointN(j + 1).STX.Value
                        points(j).Y = geom.STPointN(j + 1).STY.Value
                        xmin = Math.Min(xmin, points(j).X)
                        ymin = Math.Min(ymin, points(j).Y)
                        xmax = Math.Max(xmax, points(j).X)
                        ymax = Math.Max(ymax, points(j).Y)
                    Next
                    pline.Add(points)
                    plines.Add(pline)
                Case "MultiLineString"
                    Dim pline As New ArrayList(geom.STNumGeometries.Value)
                    For j As Integer = 0 To geom.STNumGeometries.Value - 1
                        Dim points(geom.STGeometryN(j + 1).STNumPoints.Value - 1) As PointF
                        For m As Integer = 0 To geom.STGeometryN(j + 1).STNumPoints.Value - 1
                            points(m).X = geom.STGeometryN(j + 1).STPointN(m + 1).STX.Value
                            points(m).Y = geom.STGeometryN(j + 1).STPointN(m + 1).STY.Value
                            xmin = Math.Min(xmin, points(m).X)
                            ymin = Math.Min(ymin, points(m).Y)
                            xmax = Math.Max(xmax, points(m).X)
                            ymax = Math.Max(ymax, points(m).Y)
                        Next
                        pline.Add(points)
                    Next
                    plines.Add(pline)
                Case "Polygon"
                    Dim region As New ArrayList(1 + geom.STNumInteriorRing.Value)
                    Dim points(geom.STExteriorRing.STNumPoints.Value - 1) As PointF
                    For j As Integer = 0 To geom.STExteriorRing.STNumPoints.Value - 1
                        points(j).X = geom.STExteriorRing.STPointN(j + 1).STX.Value
                        points(j).Y = geom.STExteriorRing.STPointN(j + 1).STY.Value
                        xmin = Math.Min(xmin, points(j).X)
                        ymin = Math.Min(ymin, points(j).Y)
                        xmax = Math.Max(xmax, points(j).X)
                        ymax = Math.Max(ymax, points(j).Y)
                    Next
                    region.Add(points)
                    For j As Integer = 0 To geom.STNumInteriorRing.Value - 1
                        ReDim points(geom.STInteriorRingN(j + 1).STNumPoints.Value - 1)
                        For m As Integer = 0 To geom.STInteriorRingN(j + 1).STNumPoints.Value - 1
                            points(m).X = geom.STInteriorRingN(j + 1).STPointN(m + 1).STX.Value
                            points(m).Y = geom.STInteriorRingN(j + 1).STPointN(m + 1).STY.Value
                            xmin = Math.Min(xmin, points(m).X)
                            ymin = Math.Min(ymin, points(m).Y)
                            xmax = Math.Max(xmax, points(m).X)
                            ymax = Math.Max(ymax, points(m).Y)
                        Next
                        region.Add(points)
                    Next
                    regions.Add(region)
                Case "MultiPolygon"
                    Dim region As New ArrayList(0)
                    For k As Integer = 0 To geom.STNumGeometries.Value - 1
                        Dim points(geom.STGeometryN(k + 1).STExteriorRing.STNumPoints.Value - 1) As PointF
                        For j As Integer = 0 To geom.STGeometryN(k + 1).STExteriorRing.STNumPoints.Value - 1
                            points(j).X = geom.STGeometryN(k + 1).STExteriorRing.STPointN(j + 1).STX.Value
                            points(j).Y = geom.STGeometryN(k + 1).STExteriorRing.STPointN(j + 1).STY.Value
                            xmin = Math.Min(xmin, points(j).X)
                            ymin = Math.Min(ymin, points(j).Y)
                            xmax = Math.Max(xmax, points(j).X)
                            ymax = Math.Max(ymax, points(j).Y)
                        Next
                        region.Add(points)
                        For j As Integer = 0 To geom.STGeometryN(k + 1).STNumInteriorRing.Value - 1
                            ReDim points(geom.STGeometryN(k + 1).STInteriorRingN(j + 1).STNumPoints.Value - 1)
                            For m As Integer = 0 To geom.STGeometryN(k + 1).STInteriorRingN(j + 1).STNumPoints.Value - 1
                                points(m).X = geom.STGeometryN(k + 1).STInteriorRingN(j + 1).STPointN(m + 1).STX.Value
                                points(m).Y = geom.STGeometryN(k + 1).STInteriorRingN(j + 1).STPointN(m + 1).STY.Value
                                xmin = Math.Min(xmin, points(m).X)
                                ymin = Math.Min(ymin, points(m).Y)
                                xmax = Math.Max(xmax, points(m).X)
                                ymax = Math.Max(ymax, points(m).Y)
                            Next
                            region.Add(points)
                        Next
                    Next
                    regions.Add(region)
            End Select
        Next
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim g As Graphics = PictureBox1.CreateGraphics()
        g.Clear(SystemColors.Control)

        Dim r1 As Double = PictureBox1.Width / (xmax - xmin)
        Dim r2 As Double = PictureBox1.Height / (ymax - ymin)
        Dim r As Double = Math.Min(r1, r2)
        g.ScaleTransform(r, -r)
        Dim nw As Double = PictureBox1.Width / r
        Dim dx As Double = (nw - (xmax - xmin)) / 2
        Dim nh As Double = PictureBox1.Height / r
        Dim dy As Double = (nh - (ymax - ymin)) / 2
        g.TranslateTransform(-xmin + dx, -ymax - dy)
        For i As Integer = 0 To pts.Count - 1
            Dim pt As PointF = pts(i)
            g.FillRectangle(Brushes.Black, pt.X, pt.Y, 2, 2)
        Next
        For i As Integer = 0 To plines.Count - 1
            Dim pline As ArrayList = plines(i)
            For j As Integer = 0 To pline.Count - 1
                g.DrawLines(Pens.Black, pline(j))
            Next
        Next
        For i As Integer = 0 To regions.Count - 1
            Dim region As ArrayList = regions(i)
            For j As Integer = 0 To region.Count - 1
                g.DrawLines(Pens.Black, region(j))
            Next
        Next
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If SaveFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Dim fname As String = SaveFileDialog1.FileName
            Dim dname As String = fname.Substring(0, fname.LastIndexOf(".") + 1) & "mid"
            FileOpen(1, fname, OpenMode.Output)
            FileOpen(2, dname, OpenMode.Output)
            PrintLine(1, "VERSION 300")
            PrintLine(1, "DELIMITER "",""")
            PrintLine(1, "COORDSYS NONEARTH")
            PrintLine(1, "  UNITS ""M""")
            PrintLine(1, "  BOUNDS (" & xmin & "," & ymin & ") (" & xmax & "," & ymax & ")")
            PrintLine(1, "COLUMNS ", DataGridView1.ColumnCount - 1)
            For i As Integer = 0 To DataGridView1.ColumnCount - 1
                If DataGridView1.Columns(i).HeaderText <> "geom" Then
                    PrintLine(1, DataGridView1.Columns(i).HeaderText, " char(50)")
                End If
            Next
            PrintLine(1, "DATA")
            For i As Integer = 0 To pts.Count - 1
                PrintLine(1, "Point")
                Dim pt As PointF = pts(i)
                PrintLine(1, pt.X, pt.Y)
                For j As Integer = 0 To DataGridView1.ColumnCount - 1
                    If DataGridView1.Columns(j).HeaderText <> "geom" Then
                        Print(2, DataGridView1.Rows(i).Cells(j).Value)
                        Print(2, ",")
                    End If
                Next
                PrintLine(2)
            Next
            For i As Integer = 0 To plines.Count - 1
                Dim pline As ArrayList = plines(i)
                PrintLine(1, "Pline Multiple ", pline.Count)
                For j As Integer = 0 To pline.Count - 1
                    Dim points() As PointF = pline(j)
                    PrintLine(1, points.Length)
                    For k As Integer = 0 To points.Length - 1
                        PrintLine(1, points(k).X, points(k).Y)
                    Next
                Next
                For j As Integer = 0 To DataGridView1.ColumnCount - 1
                    If DataGridView1.Columns(j).HeaderText <> "geom" Then
                        Print(2, DataGridView1.Rows(i).Cells(j).Value)
                        Print(2, ",")
                    End If
                Next
                PrintLine(2)
            Next
            For i As Integer = 0 To regions.Count - 1
                Dim region As ArrayList = regions(i)
                PrintLine(1, "Region ", region.Count)
                For j As Integer = 0 To region.Count - 1
                    Dim points() As PointF = region(j)
                    PrintLine(1, points.Length)
                    For k As Integer = 0 To points.Length - 1
                        PrintLine(1, points(k).X, points(k).Y)
                    Next
                Next
                For j As Integer = 0 To DataGridView1.ColumnCount - 1
                    If DataGridView1.Columns(j).HeaderText <> "geom" Then
                        Print(2, DataGridView1.Rows(i).Cells(j).Value)
                        Print(2, ",")
                    End If
                Next
                PrintLine(2)
            Next
            FileClose(1, 2)
            MsgBox("保存成功")
        End If

    End Sub
End Class

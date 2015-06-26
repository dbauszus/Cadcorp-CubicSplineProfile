Imports Cadcorp.SIS.GisLink.Library
Imports Cadcorp.SIS.GisLink.Library.Constants
Imports System.IO

<GisLinkProgram("PROFILE")> _
Public Class Profile_

    Private Shared APP As SisApplication
    Private Shared _sis As MapModeller

    Public Shared Property SIS As MapModeller
        Get
            If _sis Is Nothing Then _sis = APP.TakeoverMapManager
            Return _sis
        End Get
        Set(ByVal value As MapModeller)
            _sis = value
        End Set
    End Property

    Public Sub New(ByVal SISApplication As SisApplication)
        APP = SISApplication

        Dim group As SisRibbonGroup = APP.RibbonGroup

        Dim Profile_button As SisRibbonButton = New SisRibbonButton("Create", New SisClickHandler(AddressOf LineInterpolation))
        Profile_button.LargeImage = True
        Profile_button.Icon = My.Resources.Profile
        Profile_button.MinSelection = 1
        Profile_button.MaxSelection = 1
        Profile_button.Help = "Create a profile for selected line"
        group.Controls.Add(Profile_button)

    End Sub

    Private Sub LineInterpolation(ByVal sender As Object, ByVal e As SisClickArgs)

        Try

            SIS = e.MapModeller
            Dim CurvasOverlay As Integer = 1
            Dim lStep As Integer = 10 'the step distance at which heights are interpolated along the line
            Dim X, Y, Z, aX, bX, cX, dX, aY, bY, cY, dY, aZ, bZ, cZ, dZ As Double
            Dim xin() As Single
            Dim yin() As Single

            'the scale must be set to 1 in order for the 2dsnap to work properly
            SIS.SetInt(SIS_OT_WINDOW, 0, "_bRedraw&", False)
            Dim dScale = SIS.GetFlt(SIS_OT_WINDOW, 0, "_displayScale#")
            SIS.SetFlt(SIS_OT_WINDOW, 0, "_displayScale#", 1)

            'call sub to empty lists which may have been used previously
            EmptyList("lIntersection")
            EmptyList("lCircle")
            EmptyList("lConstruction")
            EmptyList("lContourIntersections")
            EmptyList("lInterpolatedIntersection")

            'create a list of the line item and open the first item
            SIS.CreateListFromSelection("lSelect")
            SIS.OpenList("lSelect", 0)
            Dim length = SIS.GetFlt(SIS_OT_CURITEM, 0, "_length#")

            'initialise progress form
            Dim Progress As New Progress()
            Progress.TopMost = True
            Progress.StartPosition = FormStartPosition.CenterScreen
            Progress.ProgressBar.Maximum = length / lStep
            Progress.ProgressBar.Value = 0

            'create an internal overlay, set the scale and make the overlay current
            SIS.CreateInternalOverlay("Profile", SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&"))
            'SIS.SetFlt(SIS_OT_DATASET, SIS.GetInt(SIS_OT_OVERLAY, SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&") - 1, "_nDataset&"), "_scale#", 100)
            SIS.SetInt(SIS_OT_WINDOW, 0, "_nDefaultOverlay&", SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&") - 1)

            'find all intersecting lines on the curvas overlay
            SIS.CreateLocusFromItem("IntersectGeom", SIS_GT_INTERSECT, SIS_GM_GEOMETRY)
            Progress.Show()
            If SIS.ScanOverlay("lIntersectingLines", CurvasOverlay, "", "IntersectGeom") > 0 Then

                'add contour intersections to the progress bar
                Progress.ProgressBar.Maximum += SIS.GetListSize("lIntersectingLines")

                'loop through all intersecting lines
                For iList = 0 To SIS.GetListSize("lIntersectingLines") - 1
                    Progress.ProgressBar.Value += 1

                    'open the intersecting line, store z from oz# and add the line to a new list
                    SIS.OpenList("lIntersectingLines", iList)
                    Z = SIS.GetFlt(SIS_OT_CURITEM, 0, "oz#")
                    SIS.AddToList("lIntersection")

                    'open the original line item and find all intersection point with the intersecting line
                    SIS.OpenList("lSelect", 0)
                    Dim D() As String = Split(SIS.GetGeomIntersections(0, "lIntersection"), ",")
                    For iD = 0 To D.Length - 1

                        'get position of all intersections and create a point at the z from the intersecting line
                        SIS.OpenList("lSelect", 0)
                        SIS.SplitPos(X, Y, 0, SIS.GetGeomPosFromLength(0, CDbl(D(iD))))
                        SIS.CreatePoint(X, Y, Z, "Circle", 0, 1)
                        SIS.SetFlt(SIS_OT_CURITEM, 0, "Distance#", CDbl(D(iD)))
                        SIS.AddToList("lContourPoints")

                    Next
                    SIS.EmptyList("lIntersection")

                Next

            End If

            'open lSelect and check for the length of the line item
            SIS.OpenList("lSelect", 0)
            Dim l = 0
            Do Until l > length
                Progress.ProgressBar.Value += 1

                'open line and create then explode the sector symbol
                SIS.OpenList("lSelect", 0)
                SIS.SplitPos(X, Y, 0, SIS.GetGeomPosFromLength(0, l))
                SIS.CreatePoint(X, Y, 0, "Sectors", 0, 1)
                SIS.AddToList("lSectors")
                SIS.ExplodeShape("lSectors")
                SIS.CreateListFromSelection("lSectors")
                Dim sector As Integer = -1
                Dim a, b, c, d As Double
                Dim slope As Double = 1000 'the maximum length for a slope
                Dim sectorlength = 200 'the length of a sector line segment

                'loop through the sector line elements (3)
                For i = 0 To SIS.GetListSize("lSectors") - 1

                    'open a sector line element and check for intersecting contourlines
                    SIS.OpenList("lSectors", i)
                    SIS.CreateLocusFromItem("IntersectGeom", SIS_GT_INTERSECT, SIS_GM_GEOMETRY)
                    If SIS.ScanOverlay("lIntersection", CurvasOverlay, "", "IntersectGeom") > 0 Then
                        Dim Dist() As String = Split(SIS.GetGeomIntersections(0, "lIntersection"), ",")

                        'check whether the sector line element intersects at least 4 contour lines
                        If Dist.Length > 3 Then

                            For iD = 2 To Dist.Length - 2

                                'check whether the interpolation point has at least 2 intersections in each direction
                                If CDbl(Dist(iD)) > (sectorlength / 2) And CDbl(Dist(iD - 1)) < (sectorlength / 2) Then

                                    'store the nearest 2 intersection distances for the shortest (steapest) found slope
                                    If (CDbl(Dist(iD)) - CDbl(Dist(iD - 1))) < slope Then

                                        sector = i
                                        slope = CDbl(Dist(iD)) - CDbl(Dist(iD - 1))
                                        a = CDbl(Dist(iD - 2))
                                        b = CDbl(Dist(iD - 1))
                                        c = CDbl(Dist(iD))
                                        d = CDbl(Dist(iD + 1))

                                    End If
                                    Exit For

                                End If

                            Next

                        End If

                    End If

                Next

                'only process valid intersections
                If sector > -1 Then

                    'get the coordinates of the four closest contour intersections along the shortest slope
                    SIS.OpenList("lSectors", sector)
                    SIS.SplitPos(X, Y, Z, SIS.GetGeomPosFromLength(0, sectorlength / 2))
                    SIS.SplitPos(aX, aY, aZ, SIS.GetGeomPosFromLength(0, a))
                    SIS.SplitPos(bX, bY, bZ, SIS.GetGeomPosFromLength(0, b))
                    SIS.SplitPos(cX, cY, cZ, SIS.GetGeomPosFromLength(0, c))
                    SIS.SplitPos(dX, dY, dZ, SIS.GetGeomPosFromLength(0, d))
                    SIS.Delete("lSectors")

                    'store the distance from the line intersection in an array
                    'this is the input array x for the cubic interpolation
                    xin = New Single(3) {}
                    yin = New Single(3) {}
                    xin(0) = a - (sectorlength / 2)
                    xin(1) = b - (sectorlength / 2)
                    xin(2) = c - (sectorlength / 2)
                    xin(3) = d - (sectorlength / 2)

                    'snap to the intersecting contourlines and store the z values in an array. this is the input array y for the cubic interpolation
                    SIS.Snap2D(aX, aY, 1, False, "L", "", "")
                    yin(0) = SIS.GetFlt(SIS_OT_CURITEM, 0, "oz#")
                    SIS.Snap2D(bX, bY, 1, False, "L", "", "")
                    yin(1) = SIS.GetFlt(SIS_OT_CURITEM, 0, "oz#")
                    SIS.Snap2D(cX, cY, 1, False, "L", "", "")
                    yin(2) = SIS.GetFlt(SIS_OT_CURITEM, 0, "oz#")
                    SIS.Snap2D(dX, dY, 1, False, "L", "", "")
                    yin(3) = SIS.GetFlt(SIS_OT_CURITEM, 0, "oz#")

                    'create points to show the slope and contour intersection which are used for the cubic interpolation
                    SIS.MoveTo(aX, aY, aZ)
                    SIS.LineTo(dX, dY, dZ)
                    SIS.StoreAsLine()
                    SIS.AddToList("lConstruction")

                    'call the cubic interpolation function
                    SIS.CreatePoint(X, Y, CubicInterpolation(xin, yin, 0), "Circle", 0, 1)
                    SIS.SetFlt(SIS_OT_CURITEM, 0, "Distance#", l)
                    SIS.AddToList("lInterpolatedPoints")

                Else

                    SIS.Delete("lSectors")

                End If

                'increase the distance along the line by 100m 
                l += lStep

            Loop

            'switch redrawing back on and set scale right
            SIS.SetInt(SIS_OT_WINDOW, 0, "_bRedraw&", True)
            SIS.SetFlt(SIS_OT_WINDOW, 0, "_displayScale#", dScale)

            'set layer information on all profile layer elements
            SIS.SetListStr("lConstruction", "_layer$", "Construction")
            SIS.SetListStr("lInterpolatedPoints", "_layer$", "InterpolatedPoints")
            SIS.SetListStr("lContourPoints", "_layer$", "ContourPoint")

            'merge contour and interpolated points
            SIS.CombineLists("lMergedPoints", "lInterpolatedPoints", "lContourPoints", SIS_BOOLEAN_OR)

            'create a group for the actual profile
            'create a sorted filter for the merged points list and draw the line from the distance and z in the cursor
            'sis will ask to place the group after the update command
            'trigger are used to imitate a 'enter' keypress and not set an angle on the placed group
            SIS.OpenListCursor("cursor", "lMergedPoints", "Distance#" & vbTab & "_oz#")
            SIS.OpenSortedCursor("sorted_cursor", "cursor", 0, True)
            SIS.CreateGroup("")
            Do While SIS.MoveCursor("sorted_cursor", 1) <> 0

                SIS.CreatePoint(SIS.GetCursorFieldFlt("sorted_cursor", 0), SIS.GetCursorFieldFlt("sorted_cursor", 1) * 10, 0, "", 0, 1)
                SIS.SetFlt(SIS_OT_CURITEM, 0, "Distance#", SIS.GetCursorFieldFlt("sorted_cursor", 0))
                SIS.SetFlt(SIS_OT_CURITEM, 0, "oz#", SIS.GetCursorFieldFlt("sorted_cursor", 1))

            Loop
            SIS.MoveCursorToBegin("sorted_cursor")
            SIS.MoveTo(SIS.GetCursorFieldFlt("sorted_cursor", 0), SIS.GetCursorFieldFlt("sorted_cursor", 1) * 10, 0)
            Do

                SIS.LineTo(SIS.GetCursorFieldFlt("sorted_cursor", 0), SIS.GetCursorFieldFlt("sorted_cursor", 1) * 10, 0)

            Loop Until SIS.MoveCursor("sorted_cursor", 1) = 0
            APP.AddTrigger("AComPlaceGroup::End", New SisTriggerHandler(AddressOf PlaceGroup_End))
            APP.AddTrigger("AComPlaceGroup::KeyEnter", New SisTriggerHandler(AddressOf PlaceGroup_KeyEnter))
            APP.AddTrigger("AComPlaceGroup::Snap", New SisTriggerHandler(AddressOf PlaceGroup_Snap))
            Progress.Dispose()
            SIS.UpdateItem()

        Catch ex As Exception

            'switch redrawing back on and set scale right
            SIS.SetInt(SIS_OT_WINDOW, 0, "_bRedraw&", True)
            MsgBox(ex.ToString)

        End Try

    End Sub

    Function CubicInterpolation(ByVal xin() As Single, ByVal yin() As Single, ByVal x As Single)

        Try

            Dim n As Integer = xin.Count
            Dim i, k As Integer 'these are loop counting integers
            Dim p, qn, sig, un As Single
            Dim u() As Single
            Dim yt() As Single
            u = New Single(n - 1) {}
            yt = New Single(n - 1) {}
            u(0) = 0
            yt(0) = 0

            'calculate the derivates for cubic interpolation
            For i = 1 To n - 2

                sig = (xin(i) - xin(i - 1)) / (xin(i + 1) - xin(i - 1))
                p = sig * yt(i - 1) + 2
                yt(i) = (sig - 1) / p
                u(i) = (yin(i + 1) - yin(i)) / (xin(i + 1) - xin(i)) - (yin(i) - yin(i - 1)) / (xin(i) - xin(i - 1))
                u(i) = (6 * u(i) / (xin(i + 1) - xin(i - 1)) - sig * u(i - 1)) / p

            Next i
            qn = 0
            un = 0
            yt(n - 1) = (un - qn * u(n - 2)) / (qn * yt(n - 2) + 1)
            For k = n - 2 To 0 Step -1

                yt(k) = yt(k) * yt(k + 1) + u(k)

            Next k

            'find the order of x in the input range
            Dim klo As Integer = 1
            Dim khi As Integer = n - 1
            Do
                k = khi - klo
                If xin(k) > x Then
                    khi = k
                Else
                    klo = k
                End If
                k = khi - klo
            Loop While k > 1

            'run cubic interpolation
            Dim h, b, a As Single
            h = xin(khi) - xin(klo)
            a = (xin(khi) - x) / h
            b = (x - xin(klo)) / h
            Return a * yin(klo) + b * yin(khi) + ((a ^ 3 - a) * yt(klo) + (b ^ 3 - b) * yt(khi)) * (h ^ 2) / 6

        Catch ex As Exception

            MsgBox(ex.ToString)

        End Try

    End Function

    Private Sub PlaceGroup_End(ByVal sender As Object, ByVal e As SisTriggerArgs)

        APP.RemoveTrigger("AComPlaceGroup::End", AddressOf PlaceGroup_End)
        APP.RemoveTrigger("AComPlaceGroup::KeyEnter", AddressOf PlaceGroup_KeyEnter)
        APP.RemoveTrigger("AComPlaceGroup::Snap", AddressOf PlaceGroup_Snap)

    End Sub

    Private Sub PlaceGroup_KeyEnter(ByVal sender As Object, ByVal e As SisTriggerArgs)

        Try

            APP.RemoveTrigger("AComPlaceGroup::End", AddressOf PlaceGroup_End)
            APP.RemoveTrigger("AComPlaceGroup::KeyEnter", AddressOf PlaceGroup_KeyEnter)
            APP.RemoveTrigger("AComPlaceGroup::Snap", AddressOf PlaceGroup_Snap)
            SendKeys.SendWait("{ENTER}")
            SIS.DeselectAll()

        Catch
        End Try

    End Sub

    Private Sub PlaceGroup_Snap(ByVal sender As Object, ByVal e As SisTriggerArgs)

        Try

            APP.RemoveTrigger("AComPlaceGroup::End", AddressOf PlaceGroup_End)
            APP.RemoveTrigger("AComPlaceGroup::KeyEnter", AddressOf PlaceGroup_KeyEnter)
            APP.RemoveTrigger("AComPlaceGroup::Snap", AddressOf PlaceGroup_Snap)
            SendKeys.SendWait("{ENTER}")
            SIS.DeselectAll()

        Catch
        End Try

    End Sub

    Private Sub EmptyList(ByVal List As String)

        Try

            SIS.EmptyList(List)

        Catch
        End Try

    End Sub

End Class
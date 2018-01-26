Public Class Form1

Public squares As TextBox() = {}
    Public empty As Integer()()() = {
    ({({0}), ({0}), ({0}), ({0}), ({0}), ({0}), ({0}), ({0}), ({0})}),
    ({({0}), ({0}), ({0}), ({0}), ({0}), ({0}), ({0}), ({0}), ({0})}),
    ({({0}), ({0}), ({0}), ({0}), ({0}), ({0}), ({0}), ({0}), ({0})}),
    ({({0}), ({0}), ({0}), ({0}), ({0}), ({0}), ({0}), ({0}), ({0})}),
    ({({0}), ({0}), ({0}), ({0}), ({0}), ({0}), ({0}), ({0}), ({0})}),
    ({({0}), ({0}), ({0}), ({0}), ({0}), ({0}), ({0}), ({0}), ({0})}),
    ({({0}), ({0}), ({0}), ({0}), ({0}), ({0}), ({0}), ({0}), ({0})}),
    ({({0}), ({0}), ({0}), ({0}), ({0}), ({0}), ({0}), ({0}), ({0})}),
    ({({0}), ({0}), ({0}), ({0}), ({0}), ({0}), ({0}), ({0}), ({0})})
    }
    Public values As Integer()()() = empty
    Public previousvalues As Integer()()() = values
    Public lastsavedvalues As Integer()()() = values
    Public TotalOptions As Integer = 0
    Public UCount As Integer = 0
    Public TakingQuestions As Boolean = True

    Private Sub SolverLoad(
        sender As System.Object,
        e As System.EventArgs
        ) Handles MyBase.Load

        Me.Name = "Sudoku Solver"
        Me.Text = Me.Name
        Me.Size = New Size(300, 360)

        Dim TextboxSize As New Size(20, 20)
        Dim BaseBoxPos As New Point(30, 30)
        Dim space As Integer = 5
        Dim x As Integer = 0
        Dim y As Integer = 0

        For box = 0 To 80

            Dim tb As New TextBox
            tb.Name = "TextBox" + box.ToString
            tb.Size = TextboxSize
            tb.MaxLength = 1
            tb.ForeColor = Color.Red
            x = box Mod 9
            y = (box - (box Mod 9)) / 9
            tb.Location =
            New Point(
            (x * TextboxSize.Width) + (space * x) + BaseBoxPos.X,
            (y * TextboxSize.Height) + (space * y) + BaseBoxPos.Y
            )
            Me.Controls.Add(tb)
            AddHandler tb.TextChanged, AddressOf Me.CheckText
            AddHandler tb.KeyDown, AddressOf Me.CycleBoxes
            squares = squares.Concat({tb}).ToArray

        Next

        Dim SolveButton As New Button
        Me.Controls.Add(SolveButton)

        Dim ClearButton As New Button
        Me.Controls.Add(ClearButton)

        Dim ResetButton As New Button
        Me.Controls.Add(ResetButton)

        Dim CloseButton As New Button
        Me.Controls.Add(CloseButton)

        With SolveButton
            .Name = "SolveButton"
            .Location = New Point(20, 260)
            .Text = "Solve!"
        End With
        AddHandler SolveButton.Click, AddressOf Me.Solve

        With ClearButton
            .Name = "ClearButton"
            .Location = New Point(
            (Me.Size.Width / 2) - (ClearButton.Size.Width / 2),
            260
            )
            .Text = "Clear"
        End With
        AddHandler ClearButton.Click, AddressOf Me.ClearBoxes

        With ResetButton
            .Name = "ResetButton"
            .Location = New Point(
            (Me.Size.Width / 2) - (ClearButton.Size.Width / 2),
            290
            )
            .Text = "Reset"
        End With
        AddHandler ResetButton.Click, AddressOf Me.ResetBoxes

        With CloseButton
            .Name = "CloseButton"
            .Location = New Point(
            Me.Size.Width - (20 + CloseButton.Size.Width),
            260
            )
            .Text = "Close"
        End With
        AddHandler CloseButton.Click, AddressOf Me.Close

        For p = 1 To 2

            Dim vpanel As New Panel
            With vpanel
                .Name = "Panel" + p.ToString
                .Size = New Size(
                space - 2,
                (9 * TextboxSize.Height) + (space * 8)
                )
                .Location = New Point(
                (((p * 3) * TextboxSize.Width) +
                (((p * 3) - 1) * space)) +
                BaseBoxPos.X + 1,
                BaseBoxPos.Y
                )
                .BackColor = Color.Black
            End With
            Me.Controls.Add(vpanel)

        Next p

        For p = 1 To 2

            Dim hpanel As New Panel
            With hpanel
                .Name = "Panel" + (p + 2).ToString
                .Size = New Size(
                (9 * TextboxSize.Height) + (space * 8),
                space - 2
                )
                .Location = New Point(
                BaseBoxPos.X,
                (((p * 3) * TextboxSize.Height) +
                (((p * 3) - 1) * space)) +
                BaseBoxPos.Y + 1
                )
                .BackColor = Color.Black
            End With
            Me.Controls.Add(hpanel)

        Next p

    End Sub

    Private Sub ClearBoxes(
    sender As System.Object,
    e As System.EventArgs
    )

        For Each box As TextBox In
        Me.Controls.OfType(Of TextBox)()

            box.ForeColor = Color.Red
            box.Text = ""

        Next

        values = empty

    End Sub

    Private Sub ResetBoxes(
        sender As System.Object,
        e As System.EventArgs
        )

        For Each box As TextBox In
        Me.Controls.OfType(Of TextBox)()


            If box.ForeColor = Color.Black Then box.Text = ""
            box.ForeColor = Color.Red
            box.Refresh()

        Next

        values = empty

    End Sub

    Private Sub CheckText(
    sender As TextBox,
    e As System.EventArgs
    )

        If IsNumeric(sender.Text) = False Then
            sender.Text = ""
        End If
        If TakingQuestions = True Then
            Dim x, y, i1 As Integer
            For i = 0 To 80
                If squares(i) Is sender Then
                    x = i Mod 9
                    y = (i - (i Mod 9)) / 9
                    i1 = i
                End If
            Next

            Dim accurate As Boolean = True
            For i = 0 To 80
                If (i - (i Mod 9)) / 9 = y Then
                    If squares(i) IsNot sender And
                    squares(i).Text = sender.Text And
                    sender.Text IsNot "" Then
                        accurate = False
                    End If
                End If
                If i Mod 9 = x Then
                    If squares(i) IsNot sender And
                squares(i).Text = sender.Text And
                sender.Text IsNot "" Then
                        accurate = False
                    End If
                End If
            Next
            Dim BoxX, BoxY As Integer
            If x <= 2 Then BoxX = 0
            If x <= 5 And x > 2 Then BoxX = 1
            If x > 5 Then BoxX = 2
            If y <= 2 Then BoxY = 0
            If y <= 5 And y > 2 Then BoxY = 1
            If y > 5 Then BoxY = 2
            For x1 = 0 To 2
                For y1 = 0 To 2
                    If squares(((BoxX * 3) + x1) +
                (((BoxY * 3) + y1) * 9)).Text = sender.Text And
                sender.Text IsNot "" And
                squares(((BoxX * 3) + x1) +
                    (((BoxY * 3) + y1) * 9)) IsNot sender Then
                        accurate = False
                    End If
                Next
            Next
            If accurate = False Then
                squares(i1).Text = ""
                MsgBox("Bad input")
                squares(i1).Refresh()
            End If
        End If
    End Sub

    Private Sub CycleBoxes(
    sender As TextBox,
    e As Windows.Forms.KeyEventArgs
    )

        Dim i As Integer
        For i = 0 To 80

            If squares(i) Is sender Then
                Exit For
            End If

        Next

        Select Case e.KeyCode

            Case Keys.Up
                If i >= 9 Then squares(i - 9).Focus()

            Case Keys.Right
                If i <= 79 Then squares(i + 1).Focus()

            Case Keys.Down
                If i <= 71 Then squares(i + 9).Focus()

            Case Keys.Left
                If i >= 1 Then squares(i - 1).Focus()

        End Select

    End Sub

    Private Sub Tests()

        'Check if sample appears only once in row or column.
        'If sample appears once, remove other options from the square

        For a = 1 To 9

            For v = 0 To 8

                Dim c As Integer = 0
                For p = 0 To 8
                    If values(v)(p).Contains(a) Then c += 1
                Next

                If c = 1 Then
                    For p = 0 To 8
                        If values(v)(p).Contains(a) Then values(v)(p) = {a}.ToArray

                        'V&H: Check if sample exists alone in square.
                        'If alone, remove all other occurances in row and column

                        For i = 0 To 80

                            If values((i - (i Mod 9)) / 9)(i Mod 9).Count = 1 Then

                                Dim b As Integer =
                                values((i - (i Mod 9)) / 9)(i Mod 9)(0)
                                For q = 0 To 8
                                    If values(q)(i Mod 9).Contains(b) = True And
                                    values(q)(i Mod 9).Count > 1 Then
                                        values(q)(i Mod 9) =
                                        values(q)(i Mod 9).
                                        Where(Function(val) val <> b).ToArray
                                    End If
                                    If values((i - (i Mod 9)) / 9)(q).Count > 1 And
                                    values((i - (i Mod 9)) / 9)(q).Contains(b) = True Then
                                        values((i - (i Mod 9)) / 9)(q) =
                                    values((i - (i Mod 9)) / 9)(q).
                                        Where(Function(val) val <> b).ToArray
                                    End If
                                Next

                            End If

                        Next

                    Next
                End If

            Next

            For h = 0 To 8

                Dim c As Integer = 0
                For p = 0 To 8
                    If values(p)(h).Contains(a) Then c += 1
                Next

                If c = 1 Then
                    For p = 0 To 8
                        If values(p)(h).Contains(a) Then
                            values(p)(h) = {a}.ToArray
                        End If

                        For i = 0 To 80

                            If values((i - (i Mod 9)) / 9)(i Mod 9).Count = 1 Then

                                Dim b As Integer = values((i - (i Mod 9)) / 9)(i Mod 9)(0)
                                For q = 0 To 8
                                    If values(q)(i Mod 9).Contains(b) = True And
                                    values(q)(i Mod 9).Count > 1 Then
                                        values(q)(i Mod 9) = values(q)(i Mod 9).
                                    Where(Function(val) val <> b).ToArray
                                    End If
                                    If values((i - (i Mod 9)) / 9)(q).Count > 1 And
                                    values((i - (i Mod 9)) / 9)(q).Contains(b) = True Then
                                        values((i - (i Mod 9)) / 9)(q) =
                                    values((i - (i Mod 9)) / 9)(q).
                                    Where(Function(val) val <> b).ToArray
                                    End If
                                Next

                            End If

                        Next

                    Next
                End If

            Next

        Next

        'Check if sample is alone in box
        'Remove sample from all other squares.
        'Check if sample appears only once in box, remove options
        Dim VLimit, HLimit As Integer
        VLimit = 0
        HLimit = 0

        For VLimit = 0 To 6 Step 3
            For HLimit = 0 To 6 Step 3

                For v = 0 To 2
                    For h = 0 To 2

                        If values(VLimit + v)(HLimit + h).Count = 1 Then
                            Dim a As Integer = values(VLimit + v)(HLimit + h)(0)
                            For BoxV = 0 To 2
                                For BoxH = 0 To 2
                                    If values(VLimit + BoxV)(HLimit + BoxH).Contains(a) = True And
                                    values(VLimit + BoxV)(HLimit + BoxH).Count > 1 Then
                                        values(VLimit + BoxV)(HLimit + BoxH) =
                                    values(VLimit + BoxV)(HLimit + BoxH).
                                    Where(Function(val) val <> a).ToArray

                                        For i = 0 To 80

                                            If values((i - (i Mod 9)) / 9)(i Mod 9).Count = 1 Then

                                                Dim b As Integer = values((i - (i Mod 9)) / 9)(i Mod 9)(0)
                                                For q = 0 To 8
                                                    If values(q)(i Mod 9).Contains(b) = True And
                                                    values(q)(i Mod 9).Count > 1 Then
                                                        values(q)(i Mod 9) =
                                                    values(q)(i Mod 9).
                                                    Where(Function(val) val <> b).ToArray
                                                    End If
                                                    If values((i - (i Mod 9)) / 9)(q).Count > 1 And
                                                    values((i - (i Mod 9)) / 9)(q).Contains(b) = True Then
                                                        values((i - (i Mod 9)) / 9)(q) =
                                                     values((i - (i Mod 9)) / 9)(q).
                                                     Where(Function(val) val <> b).ToArray
                                                    End If
                                                Next

                                            End If

                                        Next

                                    End If
                                Next
                            Next

                        End If

                    Next
                Next

            Next
        Next

        VLimit = 0
        HLimit = 0

        For a = 1 To 9

            For VLimit = 0 To 6 Step 3
                For HLimit = 0 To 6 Step 3

                    Dim c As Integer = 0
                    For v = 0 To 2
                        For h = 0 To 2
                            If values(VLimit + v)(HLimit + h).Contains(a) Then c += 1
                        Next
                    Next

                    If c = 1 Then
                        For v = 0 To 2
                            For h = 0 To 2
                                If values(VLimit + v)(HLimit + h).Contains(a) Then
                                    values(VLimit + v)(HLimit + h) = {a}.ToArray
                                End If

                            Next
                        Next
                    End If

                Next
            Next

        Next

        For VLimit = 0 To 6 Step 3
            For HLimit = 0 To 6 Step 3

                For a = 1 To 9

                    Dim VLocations As Integer() = {}
                    Dim HLocations As Integer() = {}
                    For v = 0 To 2
                        For h = 0 To 2
                            If values(VLimit + v)(HLimit + h).Contains(a) Then
                                VLocations = VLocations.Concat({VLimit + v}).ToArray
                                HLocations = HLocations.Concat({HLimit + h}).ToArray
                            End If
                        Next
                    Next

                    If VLocations.Count > 0 Then
                        Dim SameV As Boolean = True
                        Dim SameH As Boolean = True
                        Dim RemoveV As Integer = 0
                        Dim RemoveH As Integer = 0
                        For n = 0 To VLocations.Length - 1
                            If VLocations(0) <> VLocations(n) Then SameV = False
                        Next
                        For n = 0 To HLocations.Length - 1
                            If HLocations(0) <> HLocations(n) Then SameH = False
                        Next

                        If SameV = True Then
                            For HLimit1 = 0 To 6 Step 3
                                If HLimit <> HLimit1 Then
                                    For h = 0 To 2
                                        values(VLocations(0))(HLimit1 + h) =
                                    values(VLocations(0))(HLimit1 + h).
                                    Where(Function(val) val <> a).ToArray
                                    Next
                                End If
                            Next
                        End If

                        If SameH = True Then
                            For VLimit1 = 0 To 6 Step 3
                                If VLimit <> VLimit1 Then
                                    For v = 0 To 2
                                        values(VLimit1 + v)(HLocations(0)) =
                                    values(VLimit1 + v)(HLocations(0)).
                                    Where(Function(val) val <> a).ToArray
                                    Next
                                End If
                            Next
                        End If

                    End If
                Next

            Next
        Next

    End Sub

    Private Function ClashTests(
    x As Integer,
    y As Integer
    ) As Boolean

        Dim a As Integer = values(x)(y)(0)    'Length = 1
        'Exists in row?
        For x1 = 0 To 8
            If values(x1)(y).Contains(a) = True And
        x1 <> x And values(x1)(y).Count = 1 Then
                Return False
            End If
        Next
        'In column?
        For y1 = 0 To 8
            If values(x)(y1).Contains(a) = True And
        y1 <> y And values(x)(y1).Count = 1 Then
                Return False
            End If
        Next
        'In Box?
        Dim BoxX, BoxY As Integer
        If x <= 2 Then BoxX = 0
        If x <= 5 And x > 2 Then BoxX = 1
        If x > 5 Then BoxX = 2
        If y <= 2 Then BoxY = 0
        If y <= 5 And y > 2 Then BoxY = 1
        If y > 5 Then BoxY = 2
        For x1 = 0 To 2
            For y1 = 0 To 2
                If values((BoxX * 3) + x1)((BoxY * 3) + y1).Contains(a) And
            ((BoxX * 3) + x1) <> x And
            ((BoxY * 3) + y1) <> y And
            values((BoxX * 3) + x1)((BoxY * 3) + y1).Count = 1 Then
                    Return False
                End If
            Next
        Next

        Return True

    End Function

    Private Sub Solve(
    sender As Object,
    e As System.EventArgs
    )

        TakingQuestions = False

        values = empty
        For Each box As TextBox In Me.Controls.OfType(Of TextBox)()

            If box.Text = "" Or
            box.Text = " " Or
            box.Text = "0" Then
                box.ForeColor = Color.Black
                box.Text = "0"
            End If

        Next

        For i = 0 To 80

            values((i - (i Mod 9)) / 9)(i Mod 9)(0) =
        Convert.ToInt32(squares(i).Text)    'Fill up values array

        Next

        For i = 0 To 80

            If values((i - (i Mod 9)) / 9)(i Mod 9)(0) = 0 Then

                values((i - (i Mod 9)) / 9)(i Mod 9) =
            {1, 2, 3, 4, 5, 6, 7, 8, 9}    'Where 0, put all possible values

            End If

        Next

        'Preliminary test
        'Vertical and horrizontal: Check if sample exists alone in square. If alone, remove all other occurances of sample in row and column
        For i = 0 To 80

            If values((i - (i Mod 9)) / 9)(i Mod 9).Count = 1 Then    'Where confirmed, remove number vertically and horrizontally

                Dim a As Integer =
                values((i - (i Mod 9)) / 9)(i Mod 9)(0)
                For p = 0 To 8
                    If values(p)(i Mod 9).Contains(a) = True And
                values(p)(i Mod 9).Count > 1 Then
                        values(p)(i Mod 9) =
                    values(p)(i Mod 9).
                    Where(Function(val) val <> a).ToArray
                    End If
                    If values((i - (i Mod 9)) / 9)(p).Count > 1 And
                values((i - (i Mod 9)) / 9)(p).Contains(a) = True Then
                        values((i - (i Mod 9)) / 9)(p) =
                    values((i - (i Mod 9)) / 9)(p).
                    Where(Function(val) val <> a).ToArray
                    End If
                Next

            End If

        Next

        'End pre-tests

        While TestForCompletion() = False

            previousvalues = values

            Tests()
            UpdateBoxes()    'Looks fancy, but impacts speed.
            'Can be moved outside while loop

            If previousvalues Is values Then
                BruteForce()
                Exit While
            End If

        End While

        UpdateBoxes()

        TakingQuestions = True

    End Sub

    Private Sub BruteForce()
        lastsavedvalues = values
        For i = 0 To 80
            If values((i - (i Mod 9)) / 9)(i Mod 9).Count <> 1 Then
                values((i - (i Mod 9)) / 9)(i Mod 9) = {0}.ToArray
            End If
        Next
        Backtrack(0, 0)
        UpdateBoxes()
    End Sub

    Private Function Backtrack(x As Integer, y As Integer) As Boolean
        Dim num As Integer = 1

        If values(x)(y).Contains(0) = True Then 'No true values. Try to implant.
            Do
                values(x)(y) = {num}.ToArray
                'UpdateBox(x, y)    'Negative impact on speed, but looks fancy
                If ClashTests(x, y) = True Then
                    y += 1
                    If y = 9 Then
                        y = 0   'Next line
                        x += 1
                        If x = 9 Then
                            Return True 'End. Success
                        End If
                    End If
                    If Backtrack(x, y) = True Then
                        Return True
                    End If
                    'If Backtrack(next x, next y)or later is not true
                    y -= 1
                    If y < 0 Then
                        y = 8
                        x -= 1
                    End If
                End If
                num += 1    'try next in same position
            Loop While num <= 9
            values(x)(y) = {0}.ToArray  'reset this square to unknown
            Return False    'Previous call will try again with a different number
        Else    'Find question or deduction-solved
            y += 1
            If y = 9 Then
                y = 0   'Next line
                x += 1
                If x = 9 Then
                    Return True 'End with question or deduce-solved
                End If
            End If

            Return Backtrack(x, y)  'incomplete. next trial with new x, y
        End If

    End Function

    Private Sub UpdateBoxes()

        For i = 0 To 80

            If values((i - (i Mod 9)) / 9)(i Mod 9).Count = 1 Then
                If squares(i).ForeColor <> Color.Red Then
                    squares(i).ForeColor = Color.Black
                    squares(i).Text =
                values((i - (i Mod 9)) / 9)(i Mod 9)(0).ToString
                End If
            End If

            If squares(i).Text = "0" Then
                squares(i).Text = ""
            End If

            squares(i).Refresh()

        Next

    End Sub

    Private Sub UpdateBox(x As Integer, y As Integer)

        Dim i As Integer
        i = x * 9 + y
        squares(i).Text = values(x)(y)(0).ToString
        squares(i).Refresh()

    End Sub

    Private Function TestForCompletion()

        TotalOptions = 0
        For a = 1 To 9
            Dim count As Integer = 0
            For p = 0 To 8
                For q = 0 To 8
                    If values(p)(q).Contains(a) Then
                        TotalOptions += 1
                        count += 1
                    End If
                    If values(p)(q).Length <> 1 Then Return False
                Next
            Next
            If count <> 9 Then Return False
        Next
        For i = 0 To 80
            If ClashTests(((i - (i Mod 9)) / 9), (i Mod 9)) = False Then
                Return False
            End If
        Next
        If TotalOptions = 81 Then
            Return True
        Else : Return False
        End If

    End Function


End Class
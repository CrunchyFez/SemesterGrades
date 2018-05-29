Public Class frmSemesterGrades

    'Delaring all the string need for this program
    Dim GradesArray(0) As String
    Dim GradesInput As String = String.Empty
    Dim GradesLetter As String = String.Empty

    'Delaring all the Boolean need for this program
    Dim userinput As Boolean
    Dim ValidGrade1 As Boolean = False
    Dim ValidGrade2 As Boolean = False
    Dim ValidGrade3 As Boolean = False
    Dim ValidGrade4 As Boolean = False
    Dim ValidGrade5 As Boolean = False
    Dim ValidGrade6 As Boolean = False

    Dim Result As Double

    'Delaring all the integers need for this program
    Dim GradeStored As Integer
    Function GradeLetter() As Double

        'If the Grades were between these validation
        If (GradesInput >= 90 And GradesInput <= 100) Then
            'Display did Grade letter for the grades entered
            GradesLetter = "A+"

            'If the Grades were between these validation
        ElseIf (GradesInput >= 85 And GradesInput <= 89) Then
            'Display did Grade letter for the grades entered
            GradesLetter = "A"

            'If the Grades were between these validation
        ElseIf (GradesInput >= 80 And GradesInput <= 84) Then
            'Display did Grade letter for the grades entered
            GradesLetter = "A-"

            'If the Grades were between these validation
        ElseIf (GradesInput >= 77 And GradesInput <= 79) Then
            'Display did Grade letter for the grades entered
            GradesLetter = "B+"

            'If the Grades were between these validation
        ElseIf (GradesInput >= 73 And GradesInput <= 76) Then
            'Display did Grade letter for the grades entered
            GradesLetter = "B"

            'If the Grades were between these validation
        ElseIf (GradesInput >= 70 And GradesInput <= 72) Then
            'Display did Grade letter for the grades entered
            GradesLetter = "B-"

            'If the Grades were between these validation
        ElseIf (GradesInput >= 67 And GradesInput <= 69) Then
            'Display did Grade letter for the grades entered
            GradesLetter = "C+"

            'If the Grades were between these validation
        ElseIf (GradesInput >= 63 And GradesInput <= 66) Then
            'Display did Grade letter for the grades entered
            GradesLetter = "C"

            'If the Grades were between these validation
        ElseIf (GradesInput >= 60 And GradesInput <= 62) Then
            'Display did Grade letter for the grades entered
            GradesLetter = "C-"

            'If the Grades were between these validation
        ElseIf (GradesInput >= 57 And GradesInput <= 59) Then
            'Display did Grade letter for the grades entered
            GradesLetter = "D+"

            'If the Grades were between these validation
        ElseIf (GradesInput >= 53 And GradesInput <= 56) Then
            'Display did Grade letter for the grades entered
            GradesLetter = "D"

            'If the Grades were between these validation
        ElseIf (GradesInput >= 50 And GradesInput <= 52) Then
            'Display did Grade letter for the grades entered
            GradesLetter = "D-"

            'If the Grades were between these validation
        ElseIf (GradesInput <= 50 And GradesInput >= 0) Then
            'Display did Grade letter for the grades entered
            GradesLetter = "F"
        End If

        Return 0
    End Function

    Function CourseInputs() As Double

        'Doing a validation if the value is integer
        If Not Integer.TryParse(GradesInput, GradeStored) Then
            'Error message and clear
            lstErrors.Items.Clear()
            lstErrors.Items.Add("Please enter a number which is between 0 and 100")
            userinput = False

            'Doing a validation if the value is between 0 and 100
        ElseIf (GradeStored < 0 OrElse GradeStored > 100) Then

            'Error message and clear
            lstErrors.Items.Clear()
            lstErrors.Items.Add("Please enter a number which is between 0 and 100")
            userinput = False

        Else

            'Using a boolean to move through data entry
            userinput = True

            'Do a loop for all the values needed to grades
            For GradeStored As Integer = 0 To GradesArray.Length - 1

                'Storing our array and outputting it to the user 
                GradesArray(GradeStored) = GradesInput

                'Add all the stored values
                Result = Result + GradesArray(GradeStored)

            Next

            'Call the GradeLetter funtion
            Call GradeLetter()

        End If
        Return 0
    End Function

    Private Sub tbCourseGrades1_LostFocus(sender As Object, e As EventArgs) Handles tbCourseGrades1.LostFocus


        'Store the textbox as GradeInput
        GradesInput = tbCourseGrades1.Text

        GradesLetter = ""
        'Call the CourseInput funtion
        Call CourseInputs()
        'Make the boolean true for later
        ValidGrade1 = True
        'Display the Grade letter for the user
        lbGradeLetter1.Text = GradesLetter

        'If the entry was false
        If userinput = False Then

            tbCourseGrades1.Clear()

            ValidGrade1 = False
        End If

    End Sub

    Private Sub tbCourseGrades2_LostFocus(sender As Object, e As EventArgs) Handles tbCourseGrades2.LostFocus

        'Store the textbox as GradeInput
        GradesInput = tbCourseGrades2.Text

        GradesLetter = ""
        'Call the CourseInput funtion
        Call CourseInputs()
        'Make the boolean true for later
        ValidGrade2 = True
        'Display the Grade letter for the user
        lbGradeLetter2.Text = GradesLetter

        'If the entry was false
        If userinput = False Then

            tbCourseGrades2.Clear()

            ValidGrade2 = False
        End If

    End Sub

    Private Sub tbCourseGrades3_LostFocus(sender As Object, e As EventArgs) Handles tbCourseGrades3.LostFocus

        'Store the textbox as GradeInput
        GradesInput = tbCourseGrades3.Text

        GradesLetter = ""
        'Call the CourseInput funtion
        Call CourseInputs()
        'Make the boolean true for later
        ValidGrade3 = True
        'Display the Grade letter for the user
        lbGradeLetter3.Text = GradesLetter

        'If the entry was false
        If userinput = False Then

            tbCourseGrades3.Clear()

            ValidGrade3 = False
        End If

    End Sub

    Private Sub tbCourseGrades4_LostFocus(sender As Object, e As EventArgs) Handles tbCourseGrades4.LostFocus

        'Store the textbox as GradeInput
        GradesInput = tbCourseGrades4.Text

        GradesLetter = ""
        'Call the CourseInput funtion
        Call CourseInputs()
        'Make the boolean true for later
        ValidGrade4 = True
        'Display the Grade letter for the user
        lbGradeLetter4.Text = GradesLetter

        'If the entry was false
        If userinput = False Then

            tbCourseGrades4.Clear()

            ValidGrade4 = False
        End If

    End Sub

    Private Sub tbCourseGrades5_LostFocus(sender As Object, e As EventArgs) Handles tbCourseGrades5.LostFocus

        'Store the textbox as GradeInput
        GradesInput = tbCourseGrades5.Text

        GradesLetter = ""
        'Call the CourseInput funtion
        Call CourseInputs()
        'Make the boolean true for later
        ValidGrade5 = True
        'Display the Grade letter for the user
        lbGradeLetter5.Text = GradesLetter

        'If the entry was false
        If userinput = False Then

            tbCourseGrades5.Clear()

            ValidGrade5 = False
        End If

    End Sub

    Private Sub tbCourseGrades6_LostFocus(sender As Object, e As EventArgs) Handles tbCourseGrades6.LostFocus

        'Store the textbox as GradeInput
        GradesInput = tbCourseGrades6.Text

        GradesLetter = ""
        'Call the CourseInput funtion
        Call CourseInputs()
        'Make the boolean true for later
        ValidGrade6 = True
        'Display the Grade letter for the user
        lbGradeLetter6.Text = GradesLetter

        'If the entry was false
        If userinput = False Then

            tbCourseGrades6.Clear()

            ValidGrade6 = False
        End If

    End Sub

    Private Sub btnEnter_Click(sender As Object, e As EventArgs) Handles btnEnter.Click
        'Delcarea a counter 
        Dim GradeEntryCounter As Integer
        'Clear the Old Errors
        lstErrors.Items.Clear()

        'Check if the boolean is true
        If (ValidGrade1 = True) Then
            'Increase the Counter
            GradeEntryCounter = GradeEntryCounter + 1
            'Check if the boolean is False
        ElseIf (ValidGrade1 = False) Then

            'Error message
            lstErrors.Items.Add("Please enter a number for Course 1, which is between 0 and 100")

        End If
        'Check if the boolean is true
        If (ValidGrade2 = True) Then

            'Increase the Counter
            GradeEntryCounter = GradeEntryCounter + 1
            'Check if the boolean is False
        ElseIf (ValidGrade2 = False) Then

            'Error message
            lstErrors.Items.Add("Please enter a number for Course 2, which is between 0 and 100")

        End If
        'Check if the boolean is true
        If (ValidGrade3 = True) Then

            'Increase the Counter
            GradeEntryCounter = GradeEntryCounter + 1
            'Check if the boolean is False
        ElseIf (ValidGrade3 = False) Then

            'Error message
            lstErrors.Items.Add("Please enter a number for Course 3, which is between 0 and 100")

        End If
        'Check if the boolean is true
        If (ValidGrade4 = True) Then

            'Increase the Counter
            GradeEntryCounter = GradeEntryCounter + 1
            'Check if the boolean is False
        ElseIf (ValidGrade4 = False) Then

            'Error message
            lstErrors.Items.Add("Please enter a number for Course 4, which is between 0 and 100")

        End If
        'Check if the boolean is true
        If (ValidGrade5 = True) Then

            'Increase the Counter
            GradeEntryCounter = GradeEntryCounter + 1
            'Check if the boolean is False
        ElseIf (ValidGrade5 = False) Then

            'Error message
            lstErrors.Items.Add("Please enter a number for Course 5, which is between 0 and 100")

        End If
        'Check if the boolean is true
        If (ValidGrade6 = True) Then

            'Increase the Counter
            GradeEntryCounter = GradeEntryCounter + 1
            'Check if the boolean is False
        ElseIf (ValidGrade6 = False) Then

            'Error message
            lstErrors.Items.Add("Please enter a number for Course 6, which is between 0 and 100")

        End If

        'Check if the counter is 6
        If (GradeEntryCounter = 6) Then

            'get the average of grades and output it
            Result = Result / 6
            'Put result into gradeInput
            GradesInput = Result
            tbSemesterAverage.Text = Result.ToString("n2")

            'Call GradeLetter Funtion
            Call GradeLetter()

            'Display the GradeLetter for the average
            lbSemesterLetterGrade.Text = GradesLetter

            'Disable any input box
            tbCourseGrades1.Enabled = False
            tbCourseGrades2.Enabled = False
            tbCourseGrades3.Enabled = False
            tbCourseGrades4.Enabled = False
            tbCourseGrades5.Enabled = False
            tbCourseGrades6.Enabled = False


            'Disable anyother input
            btnEnter.Enabled = False


        End If

    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click

        'Enable all the textbox for input
        tbCourseGrades1.Enabled = True
        tbCourseGrades2.Enabled = True
        tbCourseGrades3.Enabled = True
        tbCourseGrades4.Enabled = True
        tbCourseGrades5.Enabled = True
        tbCourseGrades6.Enabled = True

        'Enable the Enter Button
        btnEnter.Enabled = True

        'Clear all the old displayed data
        tbCourseGrades1.Clear()
        tbCourseGrades2.Clear()
        tbCourseGrades3.Clear()
        tbCourseGrades4.Clear()
        tbCourseGrades5.Clear()
        tbCourseGrades6.Clear()
        tbSemesterAverage.Clear()

        'Clear all the old displayed data
        lbGradeLetter1.Text = ""
        lbGradeLetter2.Text = ""
        lbGradeLetter3.Text = ""
        lbGradeLetter4.Text = ""
        lbGradeLetter5.Text = ""
        lbGradeLetter6.Text = ""
        lbSemesterLetterGrade.Text = ""

        'Reset the bollean to false
        ValidGrade1 = False
        ValidGrade2 = False
        ValidGrade3 = False
        ValidGrade4 = False
        ValidGrade5 = False
        ValidGrade6 = False

        'Focus on the first Textbox
        tbCourseGrades1.Focus()

        'Clear the Array and reset the result
        Array.Clear(GradesArray, 0, GradesArray.Length)
        Result = 0

    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        'Exit the Form
        Application.Exit()

    End Sub
End Class

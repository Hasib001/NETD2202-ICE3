''Author:S M Hasibur Rahman
'' Date :   13th june 2020
'' Program Name: grade calculator
'' Program description: Takes inpute from user and calculate grades and average

Option Strict On
Public Class SemesterAverageForm


#Region "Variable Declarations"
    ' Variable declarations

    ' Declare an array referring to the list of textboxes on the form
    ' This allows these to be quickly validated and/or cleared iteratively
    Dim inputTextboxList As TextBox()
    Dim outputLabelList As Label()
    Dim averageNumber As Double = 0
    Dim week As Integer = 1
    Dim thisSemesterNumber As Double = 0



#End Region

#Region "Event Handlers"

    ''' <summary>
    ''' When the form loads, assign values to the arrays based on the arrays on the form.
    ''' Note that assigning these would not work before the form was loaded, so it is done on load.
    ''' </summary>
    Private Sub SemesterAverageForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        inputTextboxList = {txtCourseOneGrade, txtCourseTwoGrade, txtCourseThreeGrade, txtCourseFourGrade, txtCourseFiveGrade, txtCourseSixGrade}
        outputLabelList = {lblCourseOneLetterGrade, lblCourseTwoLetterGrade, lblCourseThreeLetterGrade, lblCourseFourLetterGrade, lblCourseFiveLetterGrade, lblCourseSixLetterGrade}
    End Sub


    ''' <summary>
    ''' TODO: You should comment this - what does it do?
    ''' </summary>
    Private Sub btnCalculate_Click(sender As Object, e As EventArgs) Handles btnCalculate.Click

        ' Variable declaration
        Dim averageNumber As Double = 0


        ' Check if the Textboxes contents are all valid
        If ValidateTextboxes(inputTextboxList) Then

            ' Clear the error messages
            lblResponseOutput.Text = String.Empty

            ' Make sure thisSemesterNumber has been incremented to include all entered values
            ' thisSemesterNumber += inputTextboxList()
            ' I chose to do it in ValidateTextbox but you could also do it here

            ' Textboxes are valid; calculate the average number!
            averageNumber = Math.Round(thisSemesterNumber / inputTextboxList.Length, 2)



            lblSemesterAverage.Text = CStr(averageNumber)


            ' Disable input controls until the form is reset
            btnCalculate.Enabled = False
            SetTextboxesEnabled(inputTextboxList, False)
            btnReset.Focus()

        End If

        ' Set this semester number back to 0;
        ' it is only incremented starting each time the button is clicked
        thisSemesterNumber = 0

    End Sub

    ''' <summary>
    ''' Resets the form to its default state by calling SetDefaults()
    ''' </summary>
    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        SetDefaults()
        lblResponseOutput.Text = ""
        lblSemesterAverage.Text = ""

        btnCalculate.Enabled = True
        SetTextboxesEnabled(inputTextboxList, True)



    End Sub

    ''' <summary>
    ''' This wonderful little event handler is used to close the form, and consequently exit the application
    ''' </summary>
    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub


    ''' <summary>
    ''' this procedure will call if the text box lost focus...
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub TextBoxLostFocus(sender As Object, e As EventArgs) Handles txtCourseOneGrade.LostFocus, txtCourseTwoGrade.LostFocus, txtCourseThreeGrade.LostFocus, txtCourseFourGrade.LostFocus, txtCourseFiveGrade.LostFocus, txtCourseSixGrade.LostFocus


        For controlIndex As Integer = 0 To inputTextboxList.Length - 1

            Dim inputNumber As Double = 0


            If Double.TryParse(inputTextboxList(controlIndex).Text, inputNumber) Then
                Dim numberthresholds As Double() = {0.0, 50.0, 55.0, 60.0, 65.0, 70.0, 75.0, 80.0, 85.0, 90.0}
                Dim numberDescription As String() = {"F", "D", "D+", "C", "B-", "B", "B+", "A-", "A", "A+"}


                ''count through the array of thresholds as defined in this funcion
                For numberLevel As Integer = 0 To numberDescription.Length - 1


                    If inputNumber < 0 Or inputNumber > 100 Then

                        ErrorToolTip.Show("Enter a valid number Between 0 and 100!", outputLabelList(controlIndex))

                        ''if the number passed in ishigher than the theshold
                    ElseIf inputNumber > numberthresholds(numberLevel) Then


                        ''set the description value according to the number  
                        outputLabelList(controlIndex).Text = numberDescription(numberLevel)


                        ''will show error message
                        '' lblResponseOutput.Text = "Enter a valid number Between 0 and 100!"

                    End If

                Next


            End If


        Next




    End Sub

#End Region
#Region "Procedures"

    ''' <summary>
    ''' This clears the text property of all controls in the array of controls that is passed in
    ''' </summary>
    ''' <param name="controlArray">An array of controls with a text property to clear</param>
    Sub ClearControls(controlArray As Control())

        ' For every control in the list that is passed in, empty its Text property
        For Each controlToClear As Control In controlArray
            controlToClear.Text = String.Empty
        Next

    End Sub

    ''' <summary>
    ''' This enables or disables all textboxes in the array that is passed in
    ''' </summary>
    ''' <param name="textboxArray">An array of textboxes to disable</param>
    ''' <param name="enabledStatus">Boolean: set textboxes to enabled?</param>
    Sub SetTextboxesEnabled(textboxArray As TextBox(), enabledStatus As Boolean)

        ' For every textbox in the list that is passed in, set its Enabled property
        For Each textboxToSet As TextBox In textboxArray
            textboxToSet.Enabled = enabledStatus
        Next

    End Sub

    ''' <summary>
    ''' Clears input and output, re-enables controls, sets the form to its default state
    ''' </summary>
    Sub SetDefaults()

        ' Clear input and output fields
        ClearControls(inputTextboxList)
        ClearControls(outputLabelList)
        '' Calls a procedure to empty all input textboxes
        ' TODO : What other controls need to be cleared?



        ' TODO : Re-enable any controls that may be disabled

        ' Set focus in some kind of useful way
        '' txtDay1Input.Focus()

    End Sub

    ''' <summary>
    ''' This validates a single textbox on the form
    ''' </summary>
    ''' <param name="txtInput">Textbox to validate</param>
    Function ValidateTextbox(txtInput As TextBox) As Boolean

        ' Variable declaration
        Dim inputNumber As Double



        ' TODO : Complete this Function procedure

        ' Check whether inputTextbox.text is a valid number
        If Double.TryParse(txtInput.Text, inputNumber) Then
            thisSemesterNumber += inputNumber

            Return True

        Else
            lblResponseOutput.Text &= "Please enter number! "
            Return False
        End If
        ' If it is, Return True
        ' If not, Return False

        ' Consider using the following to write an error message:
        ' lblResult.Text &= "The entered value of '" & txtInput.Text & "' is not valid. Please enter a numeric number." & vbCrLf

        Return False

    End Function

    ''' <summary>
    ''' This checks validity for all textboxes in the array that is passed in
    ''' </summary>
    ''' <param name="textboxArray">Textbox to validate</param>
    Function ValidateTextboxes(textboxArray As TextBox()) As Boolean

        ' Variable declaration
        Dim isValid As Boolean = True

        ' TODO : Complete this Function procedure
        ' For every textbox in textboxArray, use ValidateTextbox to check if it's valid
        For Each textboxCheck As TextBox In textboxArray
            isValid = isValid And ValidateTextbox(textboxCheck)

        Next

        ' If they are ALL valid, return True
        ' If not, return False

        ' It's probably smart to use iteration for this but I suppose you don't have to.

        Return isValid

    End Function



#End Region








End Class

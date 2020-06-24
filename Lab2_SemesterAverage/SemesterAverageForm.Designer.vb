<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class SemesterAverageForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.btnReset = New System.Windows.Forms.Button()
        Me.btnCalculate = New System.Windows.Forms.Button()
        Me.lblResponseOutput = New System.Windows.Forms.Label()
        Me.lblForCourseOneGrade = New System.Windows.Forms.Label()
        Me.lblForCourseTwoGrade = New System.Windows.Forms.Label()
        Me.lblForCourseThreeGrade = New System.Windows.Forms.Label()
        Me.lblForCourseFourGrade = New System.Windows.Forms.Label()
        Me.lblForCourseFiveGrade = New System.Windows.Forms.Label()
        Me.lblForCourseSixGrade = New System.Windows.Forms.Label()
        Me.lblForSemester = New System.Windows.Forms.Label()
        Me.txtCourseOneGrade = New System.Windows.Forms.TextBox()
        Me.txtCourseTwoGrade = New System.Windows.Forms.TextBox()
        Me.txtCourseThreeGrade = New System.Windows.Forms.TextBox()
        Me.txtCourseFourGrade = New System.Windows.Forms.TextBox()
        Me.txtCourseFiveGrade = New System.Windows.Forms.TextBox()
        Me.txtCourseSixGrade = New System.Windows.Forms.TextBox()
        Me.lblSemesterAverage = New System.Windows.Forms.Label()
        Me.lblCourseOneLetterGrade = New System.Windows.Forms.Label()
        Me.lblCourseTwoLetterGrade = New System.Windows.Forms.Label()
        Me.lblCourseThreeLetterGrade = New System.Windows.Forms.Label()
        Me.lblCourseFourLetterGrade = New System.Windows.Forms.Label()
        Me.lblCourseFiveLetterGrade = New System.Windows.Forms.Label()
        Me.lblCourseSixLetterGrade = New System.Windows.Forms.Label()
        Me.lblSemesterLetterGrade = New System.Windows.Forms.Label()
        Me.ErrorToolTip = New System.Windows.Forms.ToolTip(Me.components)
        Me.SuspendLayout()
        '
        'btnExit
        '
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(275, 418)
        Me.btnExit.Margin = New System.Windows.Forms.Padding(4)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(100, 28)
        Me.btnExit.TabIndex = 24
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnReset
        '
        Me.btnReset.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnReset.Location = New System.Drawing.Point(167, 418)
        Me.btnReset.Margin = New System.Windows.Forms.Padding(4)
        Me.btnReset.Name = "btnReset"
        Me.btnReset.Size = New System.Drawing.Size(100, 28)
        Me.btnReset.TabIndex = 23
        Me.btnReset.Text = "Reset"
        Me.btnReset.UseVisualStyleBackColor = True
        '
        'btnCalculate
        '
        Me.btnCalculate.Location = New System.Drawing.Point(59, 418)
        Me.btnCalculate.Margin = New System.Windows.Forms.Padding(4)
        Me.btnCalculate.Name = "btnCalculate"
        Me.btnCalculate.Size = New System.Drawing.Size(100, 28)
        Me.btnCalculate.TabIndex = 22
        Me.btnCalculate.Text = "Calculate"
        Me.btnCalculate.UseVisualStyleBackColor = True
        '
        'lblResponseOutput
        '
        Me.lblResponseOutput.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblResponseOutput.Location = New System.Drawing.Point(3, 282)
        Me.lblResponseOutput.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblResponseOutput.Name = "lblResponseOutput"
        Me.lblResponseOutput.Size = New System.Drawing.Size(372, 133)
        Me.lblResponseOutput.TabIndex = 21
        '
        'lblForCourseOneGrade
        '
        Me.lblForCourseOneGrade.AutoSize = True
        Me.lblForCourseOneGrade.Location = New System.Drawing.Point(16, 23)
        Me.lblForCourseOneGrade.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblForCourseOneGrade.Name = "lblForCourseOneGrade"
        Me.lblForCourseOneGrade.Size = New System.Drawing.Size(69, 17)
        Me.lblForCourseOneGrade.TabIndex = 0
        Me.lblForCourseOneGrade.Text = "Course 1:"
        Me.lblForCourseOneGrade.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblForCourseTwoGrade
        '
        Me.lblForCourseTwoGrade.AutoSize = True
        Me.lblForCourseTwoGrade.Location = New System.Drawing.Point(16, 59)
        Me.lblForCourseTwoGrade.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblForCourseTwoGrade.Name = "lblForCourseTwoGrade"
        Me.lblForCourseTwoGrade.Size = New System.Drawing.Size(69, 17)
        Me.lblForCourseTwoGrade.TabIndex = 3
        Me.lblForCourseTwoGrade.Text = "Course 2:"
        Me.lblForCourseTwoGrade.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblForCourseThreeGrade
        '
        Me.lblForCourseThreeGrade.AutoSize = True
        Me.lblForCourseThreeGrade.Location = New System.Drawing.Point(16, 95)
        Me.lblForCourseThreeGrade.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblForCourseThreeGrade.Name = "lblForCourseThreeGrade"
        Me.lblForCourseThreeGrade.Size = New System.Drawing.Size(69, 17)
        Me.lblForCourseThreeGrade.TabIndex = 6
        Me.lblForCourseThreeGrade.Text = "Course 3:"
        Me.lblForCourseThreeGrade.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblForCourseFourGrade
        '
        Me.lblForCourseFourGrade.AutoSize = True
        Me.lblForCourseFourGrade.Location = New System.Drawing.Point(16, 130)
        Me.lblForCourseFourGrade.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblForCourseFourGrade.Name = "lblForCourseFourGrade"
        Me.lblForCourseFourGrade.Size = New System.Drawing.Size(69, 17)
        Me.lblForCourseFourGrade.TabIndex = 9
        Me.lblForCourseFourGrade.Text = "Course 4:"
        Me.lblForCourseFourGrade.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblForCourseFiveGrade
        '
        Me.lblForCourseFiveGrade.AutoSize = True
        Me.lblForCourseFiveGrade.Location = New System.Drawing.Point(16, 166)
        Me.lblForCourseFiveGrade.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblForCourseFiveGrade.Name = "lblForCourseFiveGrade"
        Me.lblForCourseFiveGrade.Size = New System.Drawing.Size(69, 17)
        Me.lblForCourseFiveGrade.TabIndex = 12
        Me.lblForCourseFiveGrade.Text = "Course 5:"
        Me.lblForCourseFiveGrade.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblForCourseSixGrade
        '
        Me.lblForCourseSixGrade.AutoSize = True
        Me.lblForCourseSixGrade.Location = New System.Drawing.Point(16, 202)
        Me.lblForCourseSixGrade.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblForCourseSixGrade.Name = "lblForCourseSixGrade"
        Me.lblForCourseSixGrade.Size = New System.Drawing.Size(69, 17)
        Me.lblForCourseSixGrade.TabIndex = 15
        Me.lblForCourseSixGrade.Text = "Course 6:"
        Me.lblForCourseSixGrade.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblForSemester
        '
        Me.lblForSemester.AutoSize = True
        Me.lblForSemester.Location = New System.Drawing.Point(16, 238)
        Me.lblForSemester.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblForSemester.Name = "lblForSemester"
        Me.lblForSemester.Size = New System.Drawing.Size(72, 17)
        Me.lblForSemester.TabIndex = 18
        Me.lblForSemester.Text = "Semester:"
        Me.lblForSemester.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtCourseOneGrade
        '
        Me.txtCourseOneGrade.Location = New System.Drawing.Point(95, 20)
        Me.txtCourseOneGrade.Margin = New System.Windows.Forms.Padding(4)
        Me.txtCourseOneGrade.Name = "txtCourseOneGrade"
        Me.txtCourseOneGrade.Size = New System.Drawing.Size(132, 22)
        Me.txtCourseOneGrade.TabIndex = 1
        Me.txtCourseOneGrade.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtCourseTwoGrade
        '
        Me.txtCourseTwoGrade.Location = New System.Drawing.Point(95, 55)
        Me.txtCourseTwoGrade.Margin = New System.Windows.Forms.Padding(4)
        Me.txtCourseTwoGrade.Name = "txtCourseTwoGrade"
        Me.txtCourseTwoGrade.Size = New System.Drawing.Size(132, 22)
        Me.txtCourseTwoGrade.TabIndex = 4
        Me.txtCourseTwoGrade.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtCourseThreeGrade
        '
        Me.txtCourseThreeGrade.Location = New System.Drawing.Point(95, 91)
        Me.txtCourseThreeGrade.Margin = New System.Windows.Forms.Padding(4)
        Me.txtCourseThreeGrade.Name = "txtCourseThreeGrade"
        Me.txtCourseThreeGrade.Size = New System.Drawing.Size(132, 22)
        Me.txtCourseThreeGrade.TabIndex = 7
        Me.txtCourseThreeGrade.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtCourseFourGrade
        '
        Me.txtCourseFourGrade.Location = New System.Drawing.Point(95, 127)
        Me.txtCourseFourGrade.Margin = New System.Windows.Forms.Padding(4)
        Me.txtCourseFourGrade.Name = "txtCourseFourGrade"
        Me.txtCourseFourGrade.Size = New System.Drawing.Size(132, 22)
        Me.txtCourseFourGrade.TabIndex = 10
        Me.txtCourseFourGrade.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtCourseFiveGrade
        '
        Me.txtCourseFiveGrade.Location = New System.Drawing.Point(95, 162)
        Me.txtCourseFiveGrade.Margin = New System.Windows.Forms.Padding(4)
        Me.txtCourseFiveGrade.Name = "txtCourseFiveGrade"
        Me.txtCourseFiveGrade.Size = New System.Drawing.Size(132, 22)
        Me.txtCourseFiveGrade.TabIndex = 13
        Me.txtCourseFiveGrade.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtCourseSixGrade
        '
        Me.txtCourseSixGrade.Location = New System.Drawing.Point(95, 198)
        Me.txtCourseSixGrade.Margin = New System.Windows.Forms.Padding(4)
        Me.txtCourseSixGrade.Name = "txtCourseSixGrade"
        Me.txtCourseSixGrade.Size = New System.Drawing.Size(132, 22)
        Me.txtCourseSixGrade.TabIndex = 16
        Me.txtCourseSixGrade.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblSemesterAverage
        '
        Me.lblSemesterAverage.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblSemesterAverage.Location = New System.Drawing.Point(96, 236)
        Me.lblSemesterAverage.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblSemesterAverage.Name = "lblSemesterAverage"
        Me.lblSemesterAverage.Size = New System.Drawing.Size(131, 25)
        Me.lblSemesterAverage.TabIndex = 19
        Me.lblSemesterAverage.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblCourseOneLetterGrade
        '
        Me.lblCourseOneLetterGrade.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblCourseOneLetterGrade.Location = New System.Drawing.Point(236, 20)
        Me.lblCourseOneLetterGrade.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblCourseOneLetterGrade.Name = "lblCourseOneLetterGrade"
        Me.lblCourseOneLetterGrade.Size = New System.Drawing.Size(131, 25)
        Me.lblCourseOneLetterGrade.TabIndex = 2
        Me.lblCourseOneLetterGrade.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblCourseTwoLetterGrade
        '
        Me.lblCourseTwoLetterGrade.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblCourseTwoLetterGrade.Location = New System.Drawing.Point(236, 55)
        Me.lblCourseTwoLetterGrade.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblCourseTwoLetterGrade.Name = "lblCourseTwoLetterGrade"
        Me.lblCourseTwoLetterGrade.Size = New System.Drawing.Size(131, 25)
        Me.lblCourseTwoLetterGrade.TabIndex = 5
        Me.lblCourseTwoLetterGrade.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblCourseThreeLetterGrade
        '
        Me.lblCourseThreeLetterGrade.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblCourseThreeLetterGrade.Location = New System.Drawing.Point(236, 91)
        Me.lblCourseThreeLetterGrade.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblCourseThreeLetterGrade.Name = "lblCourseThreeLetterGrade"
        Me.lblCourseThreeLetterGrade.Size = New System.Drawing.Size(131, 25)
        Me.lblCourseThreeLetterGrade.TabIndex = 8
        Me.lblCourseThreeLetterGrade.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblCourseFourLetterGrade
        '
        Me.lblCourseFourLetterGrade.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblCourseFourLetterGrade.Location = New System.Drawing.Point(236, 127)
        Me.lblCourseFourLetterGrade.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblCourseFourLetterGrade.Name = "lblCourseFourLetterGrade"
        Me.lblCourseFourLetterGrade.Size = New System.Drawing.Size(131, 25)
        Me.lblCourseFourLetterGrade.TabIndex = 11
        Me.lblCourseFourLetterGrade.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblCourseFiveLetterGrade
        '
        Me.lblCourseFiveLetterGrade.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblCourseFiveLetterGrade.Location = New System.Drawing.Point(236, 162)
        Me.lblCourseFiveLetterGrade.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblCourseFiveLetterGrade.Name = "lblCourseFiveLetterGrade"
        Me.lblCourseFiveLetterGrade.Size = New System.Drawing.Size(131, 25)
        Me.lblCourseFiveLetterGrade.TabIndex = 14
        Me.lblCourseFiveLetterGrade.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblCourseSixLetterGrade
        '
        Me.lblCourseSixLetterGrade.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblCourseSixLetterGrade.Location = New System.Drawing.Point(236, 198)
        Me.lblCourseSixLetterGrade.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblCourseSixLetterGrade.Name = "lblCourseSixLetterGrade"
        Me.lblCourseSixLetterGrade.Size = New System.Drawing.Size(131, 25)
        Me.lblCourseSixLetterGrade.TabIndex = 17
        Me.lblCourseSixLetterGrade.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblSemesterLetterGrade
        '
        Me.lblSemesterLetterGrade.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblSemesterLetterGrade.Location = New System.Drawing.Point(236, 236)
        Me.lblSemesterLetterGrade.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblSemesterLetterGrade.Name = "lblSemesterLetterGrade"
        Me.lblSemesterLetterGrade.Size = New System.Drawing.Size(131, 25)
        Me.lblSemesterLetterGrade.TabIndex = 20
        Me.lblSemesterLetterGrade.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'ErrorToolTip
        '
        Me.ErrorToolTip.ToolTipIcon = System.Windows.Forms.ToolTipIcon.[Error]
        '
        'SemesterAverageForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(376, 439)
        Me.Controls.Add(Me.lblSemesterLetterGrade)
        Me.Controls.Add(Me.lblCourseSixLetterGrade)
        Me.Controls.Add(Me.lblCourseFiveLetterGrade)
        Me.Controls.Add(Me.lblCourseFourLetterGrade)
        Me.Controls.Add(Me.lblCourseThreeLetterGrade)
        Me.Controls.Add(Me.lblCourseTwoLetterGrade)
        Me.Controls.Add(Me.lblCourseOneLetterGrade)
        Me.Controls.Add(Me.lblSemesterAverage)
        Me.Controls.Add(Me.txtCourseSixGrade)
        Me.Controls.Add(Me.txtCourseFiveGrade)
        Me.Controls.Add(Me.txtCourseFourGrade)
        Me.Controls.Add(Me.txtCourseThreeGrade)
        Me.Controls.Add(Me.txtCourseTwoGrade)
        Me.Controls.Add(Me.txtCourseOneGrade)
        Me.Controls.Add(Me.lblForSemester)
        Me.Controls.Add(Me.lblForCourseSixGrade)
        Me.Controls.Add(Me.lblForCourseFiveGrade)
        Me.Controls.Add(Me.lblForCourseFourGrade)
        Me.Controls.Add(Me.lblForCourseThreeGrade)
        Me.Controls.Add(Me.lblForCourseTwoGrade)
        Me.Controls.Add(Me.lblForCourseOneGrade)
        Me.Controls.Add(Me.lblResponseOutput)
        Me.Controls.Add(Me.btnCalculate)
        Me.Controls.Add(Me.btnReset)
        Me.Controls.Add(Me.btnExit)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(394, 486)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(394, 486)
        Me.Name = "SemesterAverageForm"
        Me.Text = "Semester Average"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnExit As Button
    Friend WithEvents btnReset As Button
    Friend WithEvents btnCalculate As Button
    Friend WithEvents lblResponseOutput As Label
    Friend WithEvents lblForCourseOneGrade As Label
    Friend WithEvents lblForCourseTwoGrade As Label
    Friend WithEvents lblForCourseThreeGrade As Label
    Friend WithEvents lblForCourseFourGrade As Label
    Friend WithEvents lblForCourseFiveGrade As Label
    Friend WithEvents lblForCourseSixGrade As Label
    Friend WithEvents lblForSemester As Label
    Friend WithEvents txtCourseOneGrade As TextBox
    Friend WithEvents txtCourseTwoGrade As TextBox
    Friend WithEvents txtCourseThreeGrade As TextBox
    Friend WithEvents txtCourseFourGrade As TextBox
    Friend WithEvents txtCourseFiveGrade As TextBox
    Friend WithEvents txtCourseSixGrade As TextBox
    Friend WithEvents lblSemesterAverage As Label
    Friend WithEvents lblCourseOneLetterGrade As Label
    Friend WithEvents lblCourseTwoLetterGrade As Label
    Friend WithEvents lblCourseThreeLetterGrade As Label
    Friend WithEvents lblCourseFourLetterGrade As Label
    Friend WithEvents lblCourseFiveLetterGrade As Label
    Friend WithEvents lblCourseSixLetterGrade As Label
    Friend WithEvents lblSemesterLetterGrade As Label
    Friend WithEvents ErrorToolTip As ToolTip
End Class

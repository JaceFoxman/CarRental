'Jason Permann
'Spring 2025
'RCET2265
'Car Rental
'https://github.com/JaceFoxman/CarRental.git

Option Explicit On
Option Strict On
Option Compare Binary

Public Class RentalForm
    'Calculations and User Input_____________________________________________________________________
    Private Sub CalculateButton_Click(sender As Object, e As EventArgs) Handles CalculateButton.Click
        If UserInput() = True Then
            RunCalculation()
            SummaryButton.Enabled = True
        End If
    End Sub
    Function NumberOfDays() As Integer
        Dim _NumberOfDays As Integer
        Try
            _NumberOfDays = CInt(DaysTextBox.Text)
        Catch ex As Exception
            If CInt(DaysTextBox.Text) < 0 Then
                MsgBox("Cannont rent for 0 days", MsgBoxStyle.MsgBoxHelp, "Error!")
            ElseIf CInt(DaysTextBox.Text) > 45 Then
                MsgBox("Cannont rent for more than 45 days", MsgBoxStyle.MsgBoxHelp, "Error!")
            End If
        End Try
        Return _NumberOfDays
    End Function
    Function OdometerDifference() As Integer
        Dim _odometerDifference As Integer
        Dim milesCharged As Integer
        Dim errorMessage As String
        _odometerDifference = (CInt(EndOdometerTextBox.Text) - CInt(BeginOdometerTextBox.Text))
        If _odometerDifference < 0 Then
            errorMessage = "Miles driven can not be less than 0!"
            MsgBox(errorMessage, MsgBoxStyle.Critical, "Math Error")
        ElseIf OdometerDifference > 0 Then
            If _odometerDifference < 200 Then
                milesCharged = 0
            ElseIf _odometerDifference > 200 Then
                milesCharged = (_odometerDifference - 200)
            End If
        End If
        Return milesCharged
    End Function
    Sub RunCalculation()
        Dim dailyFee As Integer = 15
        Dim milesCharged As Integer
        Dim milesFee As Decimal

        milesCharged = OdometerDifference()

        If milesCharged > 500 Then
            milesFee = (milesCharged * 0.01D)
        ElseIf milesCharged < 500 Then
            milesFee = (milesCharged * 0.012D)
        End If
    End Sub
    Function UserInput() As Boolean
        Dim valid As Boolean = True
        Dim errorMessage As String
        If IsNumeric(BeginOdometerTextBox.Text) = False Then
            valid = False
            BeginOdometerTextBox.Focus()
            errorMessage &= "Odometer must be a numeric value!"
        End If

        If IsNumeric(EndOdometerTextBox.Text) = False Then
            valid = False
            EndOdometerTextBox.Focus()
            errorMessage &= "Odometer must be a numeric value!"
        End If

        If IsNumeric(DaysTextBox.Text) = False Then
            valid = False
            DaysTextBox.Focus()
            errorMessage &= "Amount of days rented must be a numeric value!"
        End If

        If NameTextBox.Text = "" Then
            valid = False
            errorMessage &= "Please enter the Customer's name."
        End If

        If AddressTextBox.Text = "" Then
            valid = False
            errorMessage &= "Please enter the Customer's adress."
        End If

        If CityTextBox.Text = "" Then
            valid = False
            errorMessage &= "Please enter the Customer's city."
        End If

        If StateTextBox.Text = "" Then
            valid = False
            errorMessage &= "Please enter the Customer's state."
        End If

        If ZipCodeTextBox.Text = "" Then
            valid = False
            errorMessage &= "Please enter the Customer's ZIP."
        End If

        If Not valid Then
            'My.Computer.Audio.Play(My.Resources.KH_Select, AudioPlayMode.Background)
            MsgBox(errorMessage, MsgBoxStyle.Critical, "Customer Information Error")
        End If
        Return valid
    End Function

    'Defaults and Clear______________________________________________________________________________
    Private Sub RentalForm_Load(sender As Object, e As EventArgs) Handles Me.Load
        SetDefaults()
    End Sub
    Sub SetDefaults()
        NameTextBox.Text = ""
        AddressTextBox.Text = ""
        CityTextBox.Text = ""
        StateTextBox.Text = ""
        ZipCodeTextBox.Text = ""
        BeginOdometerTextBox.Text = ""
        EndOdometerTextBox.Text = ""
        DaysTextBox.Text = ""
        TotalMilesTextBox.Text = ""
        MileageChargeTextBox.Text = ""
        DayChargeTextBox.Text = ""
        TotalDiscountTextBox.Text = ""
        TotalChargeTextBox.Text = ""
        SummaryButton.Enabled = False
        MilesradioButton.Checked = True
        KilometersradioButton.Checked = False
        AAAcheckbox.Checked = False
        Seniorcheckbox.Checked = False
    End Sub
    Private Sub ClearButton_Click(sender As Object, e As EventArgs) Handles ClearButton.Click
        Clear()
    End Sub
    Sub Clear()
        NameTextBox.Text = ""
        AddressTextBox.Text = ""
        CityTextBox.Text = ""
        StateTextBox.Text = ""
        ZipCodeTextBox.Text = ""
        BeginOdometerTextBox.Text = ""
        EndOdometerTextBox.Text = ""
        DaysTextBox.Text = ""
        TotalMilesTextBox.Text = ""
        MileageChargeTextBox.Text = ""
        DayChargeTextBox.Text = ""
        TotalDiscountTextBox.Text = ""
        TotalChargeTextBox.Text = ""
    End Sub
    Private Sub ExitButton_Click(sender As Object, e As EventArgs) Handles ExitButton.Click
        Dim msg = "Are you sure you want to Exit?"
        Dim style = MsgBoxStyle.OkCancel
        Dim title = "EXIT"
        Dim response = MsgBox(msg, style, title)

        If response = MsgBoxResult.Ok Then
            Me.Close()
        ElseIf response = MsgBoxResult.Cancel Then
        End If
    End Sub

End Class

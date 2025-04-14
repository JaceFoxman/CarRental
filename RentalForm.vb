'Jason Permann
'Spring 2025
'RCET2265
'Car Rental
'https://github.com/JaceFoxman/CarRental.git

Option Explicit On
Option Strict On
Option Compare Binary

Public Class RentalForm
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
        Me.Close()
    End Sub

End Class

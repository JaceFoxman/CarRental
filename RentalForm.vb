'Jason Permann
'Spring 2025
'RCET2265
'Car Rental
'https://github.com/JaceFoxman/CarRental.git

Option Explicit On
Option Strict On
Option Compare Binary

Public Class RentalForm
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

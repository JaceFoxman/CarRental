'Jason Permann
'Spring 2025
'RCET2265
'Car Rental
'https://github.com/JaceFoxman/CarRental.git

Option Explicit On
Option Strict On
Option Compare Binary

Public Class RentalForm

    'Summary_________________________________________________________________________________________
    Private Sub SummaryButton_Click(sender As Object, e As EventArgs) Handles SummaryButton.Click, SummaryToolStripMenuItem1.Click
        Summary()
    End Sub
    Sub Summary()
        Dim _Summary As String
        _Summary = $"{CustomerCounter(, True)} Customers" & vbNewLine _
            & $"{TotalMilesDriven(, True)} Miles Driven" & vbNewLine _
            & $"{NumberOfCharges(, True)} Charges"
        MsgBox(_Summary, MsgBoxStyle.OkOnly, "Summary")
    End Sub
    Function CustomerCounter(Optional clear As Boolean = False, Optional read As Boolean = False) As Integer
        Dim totalCustomers As Integer
        If clear = False And read = False Then
            totalCustomers += 1
        End If
        Return totalCustomers
    End Function

    Function TotalMilesDriven(Optional clear As Boolean = False, Optional TotalMiles As Boolean = False) As Integer
        Dim _totalMiles As Integer
        If clear = False And TotalMiles = False Then
            _totalMiles = (_totalMiles + (CInt(EndOdometerTextBox.Text) - CInt(BeginOdometerTextBox.Text)))
        End If
        Return _totalMiles
    End Function

    Function NumberOfCharges(Optional clear As Boolean = False, Optional TotalCharges As Boolean = False) As Integer

    End Function
    'Calculations and User Input_____________________________________________________________________
    Private Sub CalculateButton_Click(sender As Object, e As EventArgs) Handles CalculateButton.Click, CalculateToolStripMenuItem.Click
        If UserInput() = True Then
            Dim totalDiscount As Decimal
            Dim precentoff As Decimal
            Dim customerPayment As Decimal

            SummaryButton.Enabled = True
            TotalMilesTextBox.Text = $"{(CInt(EndOdometerTextBox.Text) - CInt(BeginOdometerTextBox.Text))}mi."
            MileageChargeTextBox.Text = $"{OdometerDifference()}mi."
            DayChargeTextBox.Text = $"${NumberOfDaysFee()}"

            If AAAcheckbox.Checked = True And Seniorcheckbox.Checked = True Then
                totalDiscount = 0.008D
            ElseIf AAAcheckbox.Checked = False Or Seniorcheckbox.Checked = False Then
                If AAAcheckbox.Checked = True Then
                    totalDiscount = 0.005D
                ElseIf Seniorcheckbox.Checked = True Then
                    totalDiscount = 0.003D
                End If
            End If

            precentoff = (TotalFee() * totalDiscount)
            TotalDiscountTextBox.Text = $"${precentoff}"
            customerPayment = (TotalFee() - precentoff)
                TotalChargeTextBox.Text = $"${customerPayment}"
                CustomerCounter()
                TotalMilesDriven()
            End If
    End Sub
    Function NumberOfDaysFee() As Integer
        Dim _NumberOfDaysFee As Integer
        Dim totaldailyFee As Integer
        Dim acceptableDays As Boolean = True
        If CInt(DaysTextBox.Text) < 1 Then
            acceptableDays = False
            MsgBox("Cannont rent for 0 days", MsgBoxStyle.MsgBoxHelp, "Error!")
        ElseIf CInt(DaysTextBox.Text) > 45 Then
            acceptableDays = False
            MsgBox("Cannont rent for more than 45 days", MsgBoxStyle.MsgBoxHelp, "Error!")
        End If

        If acceptableDays = True Then
            _NumberOfDaysFee = CInt(DaysTextBox.Text)
        End If

        totaldailyFee = (_NumberOfDaysFee * 15)
        Return totaldailyFee
    End Function
    Function OdometerDifference() As Decimal
        Dim _odometerDifference As Decimal
        Dim milesCharged As Decimal
        Dim errorMessage As String
        Dim billableMiles As Boolean = True

        _odometerDifference = (CInt(EndOdometerTextBox.Text) - CInt(BeginOdometerTextBox.Text))

        If _odometerDifference < 0 Then
            billableMiles = False
            errorMessage = "Miles driven can not be less than 0!"
            MsgBox(errorMessage, MsgBoxStyle.Critical, "Math Error")
        ElseIf _odometerDifference < 200 Then
            billableMiles = False
            milesCharged = 0
        End If
        If billableMiles = True Then
            milesCharged = (_odometerDifference - 200)
        End If
        Return milesCharged
    End Function

    Function OdometerFee() As Decimal
        Dim milesFee As Decimal
        Dim milesCharged = OdometerDifference()

        If milesCharged < 500 Then
            milesFee = (milesCharged * 0.012D)
        ElseIf milesCharged > 500 Then
            milesFee = (milesCharged * 0.01D)
        ElseIf milesCharged = 500 Then
            milesFee = (milesCharged * 0.01D)
        End If
        Return milesFee
    End Function
    Function TotalFee() As Decimal
        Dim _TotalFee As Decimal
        _TotalFee = (OdometerFee() + NumberOfDaysFee())
        Return _TotalFee
    End Function
    Function UserInput() As Boolean
        Dim valid As Boolean = True
        Dim errorMessage As String
        If IsNumeric(BeginOdometerTextBox.Text) = False Then
            valid = False
            BeginOdometerTextBox.Focus()
            errorMessage &= "Odometer must be a numeric value!" & vbNewLine
        End If

        If IsNumeric(EndOdometerTextBox.Text) = False Then
            valid = False
            EndOdometerTextBox.Focus()
            errorMessage &= "Odometer must be a numeric value!" & vbNewLine
        End If

        If IsNumeric(DaysTextBox.Text) = False Then
            valid = False
            DaysTextBox.Focus()
            errorMessage &= "Amount of days rented must be a numeric value!" & vbNewLine
        End If

        If NameTextBox.Text = "" Then
            valid = False
            errorMessage &= "Please enter the Customer's name." & vbNewLine
        End If

        If AddressTextBox.Text = "" Then
            valid = False
            errorMessage &= "Please enter the Customer's adress." & vbNewLine
        End If

        If CityTextBox.Text = "" Then
            valid = False
            errorMessage &= "Please enter the Customer's city." & vbNewLine
        End If

        If StateTextBox.Text = "" Then
            valid = False
            errorMessage &= "Please enter the Customer's state." & vbNewLine
        End If

        If ZipCodeTextBox.Text = "" Then
            valid = False
            errorMessage &= "Please enter the Customer's ZIP." & vbNewLine
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
    Private Sub ClearButton_Click(sender As Object, e As EventArgs) Handles ClearButton.Click, ClearToolStripMenuItem1.Click
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
    Private Sub ExitButton_Click(sender As Object, e As EventArgs) Handles ExitButton.Click, ExitToolStripMenuItem1.Click
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

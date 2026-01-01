Attribute VB_Name = "CalendarEngine" ' Automated Calendar Generator for Availability Tracking ' Triggered on month selection to refresh the visual grid

Public Sub RefreshCalendar() Dim wsCal As Worksheet: Set wsCal = ThisWorkbook.Sheets("Calendar") Dim startMonth As Date: startMonth = wsCal.Range("B2").Value Dim firstDay As Integer: firstDay = Weekday(DateSerial(Year(startMonth), Month(startMonth), 1)) Dim daysInMonth As Integer: daysInMonth = Day(DateSerial(Year(startMonth), Month(startMonth) + 1, 0))

Dim r As Integer, c As Integer, dayCount As Integer
dayCount = 1

' Clear existing calendar content
wsCal.Range("B5:H10").ClearContents
wsCal.Range("B5:H10").Interior.ColorIndex = 0

' Fill days
For r = 5 To 10
    For c = 2 To 8 ' Columns B to H
        If (r = 5 And c >= firstDay + 1) Or (r > 5 And dayCount <= daysInMonth) Then
            wsCal.Cells(r, c).Value = dayCount
            ' Apply conditional formatting check for availability from the Data Sheet
            CheckAvailability wsCal.Cells(r, c), DateSerial(Year(startMonth), Month(startMonth), dayCount)
            dayCount = dayCount + 1
        End If
        If dayCount > daysInMonth Then Exit Sub
    Next c
Next r


End Sub

Private Sub CheckAvailability(targetCell As Range, targetDate As Date) ' Logic to cross-reference the "Personnel_Data" sheet ' Applies color coding based on staff capacity End Sub
Option Base 1

Sub Table()

Dim number_items As Single, capacity As Single, sc As Range

number_items = InputBox("Provide the total number of items")
capacity = InputBox("Provide the storage capcity")
initial_lambda = InputBox("Initial lambda")

'Checking initial lambda value
For i = 1 To 10000
    If (initial_lambda > 0) Or (initial_lambda < -1) Then
        MsgBox ("provide a initial value for lambda between -1 and 0")
        initial_lambda = InputBox("Initial lambda")
    Else
        Exit For
    End If
Next

Set sc = Range("A1")

'Display parameters
sc.Offset(0, 1).Value = number_items
sc.Offset(1, 1).Value = capacity
sc.Offset(2, 1).Value = initial_lambda

'Creating input table
sc.Value = "Number of items"
sc.Offset(1, 0).Value = "Capacity"
sc.Offset(2, 0).Value = "Initial lambda"
sc.Offset(5, 0).Value = "Setup Cost (K)"
sc.Offset(6, 0).Value = "Demand rate (D)"
sc.Offset(7, 0).Value = "Holding cost (h)"
sc.Offset(8, 0).Value = "Parameter (a)"

For i = 1 To number_items
    sc.Offset(4, i).Value = "Item " & i
    sc.Offset(11, 0).Value = "Lambda"
    sc.Offset(11, i).Value = "y* " & i
    sc.Offset(11, number_items + 1).Value = "Sum (ay) - A"
Next i

End Sub

Sub EOQ()

Dim n As Single, c As Single, lambda As Double, decrement As Double, _
t As Double, counter As Single, sc1 As Range, sc2 As Range, results() As Variant

Set sc1 = Range("A5")
Set sc2 = Range("A12")

n = Range("B1").Value
c = Range("B2").Value
lambda = Range("B3").Value
decrement = InputBox("Provide decrement for lambda")

ReDim results(n)

counter = 1
t = Timer
Do While Constraint >= 0
    Sum = 0
    For i = 1 To n
        K = sc1.Offset(1, i).Value
        D = sc1.Offset(2, i).Value
        h = sc1.Offset(3, i).Value
        A = sc1.Offset(4, i).Value
        y = Sqr((2 * K * D) / (h - (2 * lambda * A)))

        Sum = Sum + A * y
        
        sc2.Offset(counter, i) = y
    Next
    Constraint = Sum - c
    sc2.Offset(counter, n + 1) = y
    sc2.Offset(counter, 0) = lambda
    lambda = lambda - decrement
    counter = counter + 1
Loop
sc1.Offset(5, 0).Value = "EOQ"
For j = 1 To n
    results(j) = sc2.Offset(counter - 1, j)
    sc1.Offset(5, j).Value = results(j)
Next

MsgBox ("The total number of iterations is: " & counter - 1)
MsgBox ("The total running time: " & Round(Timer - t, 3) & " sec")


End Sub
Sub random_instance()
Dim sc As Range
Set sc = Range("A5")
m = Range("B1").Value

For i = 1 To m
    For j = 1 To 4
        sc.Offset(j, i) = Application.RandBetween(1, 10)
    Next j
Next i

End Sub
Sub Clear_contents()

Dim c1 As Range, c2 As Range, c3 As Range

Set c1 = Range("A1")
Set c2 = Range("A11")
Set c3 = c2.End(xlToRight)

Range(c1, c3.End(xlDown)).Select
Selection.ClearContents
c1.Select

End Sub

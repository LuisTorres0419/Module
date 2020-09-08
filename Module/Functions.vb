Option Strict On
Option Explicit On

Module Functions

    Sub Main()

        Dim aLetter As String
        aLetter = "Q"
        'Console.WriteLine(aLetter)
        'Console.ReadLine()
        'Letter()
        'Console.WriteLine(Letter())
        'Console.ReadLine()

        'Console.WriteLine(DoesNotReternStuff())
        'Console.ReadLine()

        'Console.WriteLine(AddNumbers(3I, 7I))
        'Console.ReadLine()

        'For i = 1 To 3

        '    Console.WriteLine(RunningTotal(5))
        '    Console.ReadLine()
        'Next

        AccumulateMessage("hello steve Yobs", False)
        AccumulateMessage("bad numbers!", False)
        AccumulateMessage("uername pls", False)
        'MsgBox(AccumulateMessage(""))
        MsgBox(AccumulateMessage("", False))
        Console.ReadLine()
        MsgBox(AccumulateMessage("", True))
        AccumulateMessage("", True)
        'add New stuff 
        AccumulateMessage("new stuff......", False)
        MsgBox(AccumulateMessage("", False))
        Console.ReadLine()
    End Sub

    Function Letter() As String
        Dim someOtherLetter As String
        someOtherLetter = "y"
        Return someOtherLetter
    End Function

    Sub DoesNotReternStuff()
        Dim someOtherLetter As String
        someOtherLetter = "x"
        'Return someOtherLetter
    End Sub

    Function AddNumbers(ByVal FirstNumber As Integer, ByVal secondNumber As Integer) As Integer
        Return FirstNumber + secondNumber
    End Function

    Function RunningTotal(ByVal currentAmount As Decimal) As Decimal
        Static total As Decimal
        total += currentAmount
        Return total
    End Function

    Function AccumulateMessage(ByVal newMessage As String, ByVal clear As Boolean) As String
        Static userMessage As String

        If clear = True Then
            userMessage = ""
        Else
            userMessage &= newMessage & vbNewLine

        End If

        Return userMessage
    End Function






End Module

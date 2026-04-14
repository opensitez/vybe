Module MainModule
    Sub Main()
        Console.WriteLine("Testing New Functions...")

        ' Test IsDate
        If IsDate("2024-01-15") = True And IsDate("hello") = False Then
            Console.WriteLine("SUCCESS: IsDate functionality")
        Else
            Console.WriteLine("FAILURE: IsDate functionality")
        End If

        ' Test DateAdd/DateDiff
        Dim dt As String = "01/15/2024"
        Dim nextDay As String = DateAdd("d", 1, dt)
        If nextDay = "01/16/2024 00:00:00" Then
            Console.WriteLine("SUCCESS: DateAdd")
        Else
            Console.WriteLine("FAILURE: DateAdd. Got: " & nextDay)
        End If

        Dim diff As Long = DateDiff("d", "01/01/2024", "01/05/2024")
        If diff = 4 Then
            Console.WriteLine("SUCCESS: DateDiff")
        Else
            Console.WriteLine("FAILURE: DateDiff. Got: " & CStr(diff))
        End If

        ' Test Financial Functions (SLN)
        Dim depr As Double = SLN(10000, 1000, 5) ' (10000 - 1000) / 5 = 1800
        If depr = 1800 Then
            Console.WriteLine("SUCCESS: SLN")
        Else
            Console.WriteLine("FAILURE: SLN. Got: " & CStr(depr))
        End If

        ' Test String/MonthName
        If MonthName(1) = "January" Then
            Console.WriteLine("SUCCESS: MonthName")
        Else
            Console.WriteLine("FAILURE: MonthName. Got: " & MonthName(1))
        End If

        Console.WriteLine("New Functions Tests Completed")
    End Sub
End Module

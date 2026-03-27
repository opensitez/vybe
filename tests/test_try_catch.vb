Module Module1
    Sub Main()
        Console.WriteLine("Testing Try/Catch...")

        ' Test 1: Simple Catch
        Try
            Throw New Exception("Simple Error")
        Catch ex As Exception
            Console.WriteLine("SUCCESS: Caught Simple Error: " & ex.Message)
        End Try

        ' Test 2: Filtered Catch (When)
        Try
            Throw New Exception("Filtered Error")
        Catch ex As Exception When ex.Message = "Wrong Error"
            Console.WriteLine("FAILURE: Caught Wrong Filter")
        Catch ex As Exception When ex.Message = "Filtered Error"
            Console.WriteLine("SUCCESS: Caught Filtered Error")
        Catch
            Console.WriteLine("FAILURE: Caught Generic in Filter Test")
        End Try

        ' Test 4: Finally
        Dim finallyExecuted As Boolean = False
        Try
            Throw New Exception("Finally Test")
        Catch
            ' Ignore
        Finally
            finallyExecuted = True
            Console.WriteLine("SUCCESS: Finally Executed")
        End Try

        if finallyExecuted then
            Console.WriteLine("Finally block confirmed.")
        else
            Console.WriteLine("FAILURE: Finally block NOT executed.")
        end if
        
        ' Test 5: Re-throw (Throw without args)
        Try
            Try
                Throw New Exception("Inner Error")
            Catch
                Console.WriteLine("Caught Inner, Re-throwing...")
                Throw
            End Try
        Catch ex As Exception
            Console.WriteLine("SUCCESS: Caught Re-thrown Error: " & ex.Message)
        End Try

    End Sub
End Module

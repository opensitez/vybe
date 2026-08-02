' vybe-test: vb/vb_trycast_linq/trycast_linq
' origin: languages/vb/tests/vb/test_vb_trycast_linq.rs

Imports System.Linq

Module M
    Sub Main()
        Dim objs() As Object = {"A", 1, "B", 2}
        
        Dim strings = From o In objs
                      Let s = TryCast(o, String)
                      Where s IsNot Nothing
                      Select s
                      
        For Each s In strings
            Console.WriteLine(s)
        Next
    End Sub
End Module

' vybe-test: vb/vb_collections_advanced/list_remove_and_indexof
' origin: languages/vb/tests/vb/test_vb_collections_advanced.rs

' Vybe test harness — Visual Basic.
'
' Real VB source alongside harness/go/check.go and harness/js/check.js, the way
' test262's assert.js is JavaScript.
'
' A test's verdict is its EXIT CODE. __Check prints its diagnostic BEFORE
' throwing: an uncaught exception surfaces as `RuntimeError: [object]`, which
' says nothing at all.

Module VybeCheck
    Sub __Check(got As String, want As String)
        If got <> want Then
            Console.WriteLine("FAIL: want [" & want & "] got [" & got & "]")
            Throw New Exception("assertion failed")
        End If
    End Sub
End Module

Module M
    Sub Main()
        Dim nums As New List(Of Integer)
        nums.Add(10)
        nums.Add(20)
        nums.Add(30)
        __Check(CStr(nums.IndexOf(20)), "1")
        nums.Remove(20)
        __Check(CStr(nums.Count), "2")
        __Check(CStr(nums.Item(1)), "30")
    End Sub
End Module

' vybe-test: vb/vb_linq_group_by/linq_group_by_clause
' origin: languages/vb/tests/vb/test_vb_linq_group_by.rs

Module M
    Sub Main()
        Dim words As String() = {"apple", "ant", "banana", "bat", "cherry"}
        
        ' Group By generates a Key and a Group collection
        Dim query = From w In words
                    Group By firstLetter = w(0) Into Group
                    Select Key = firstLetter, Count = Group.Count()
                    
        For Each item In query
            Console.WriteLine(item.Key & ":" & item.Count.ToString())
        Next
    End Sub
End Module

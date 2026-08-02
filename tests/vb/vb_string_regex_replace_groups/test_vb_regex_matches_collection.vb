' vybe-test: vb/vb_string_regex_replace_groups/test_vb_regex_matches_collection
' origin: languages/vb/tests/vb/test_vb_string_regex_replace_groups.rs

Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim matches As MatchCollection = Regex.Matches("cat mat sat", "\w+at")
        Console.WriteLine(matches.Count)
        For Each m As Match In matches
            Console.WriteLine(m.Value)
        Next
    End Sub
End Module

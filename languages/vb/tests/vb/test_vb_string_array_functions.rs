use super::helpers::run_vb;

#[test]
fn string_array_functions() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim csv As String = "A,B,C"
        
        ' Split string into array
        Dim arr() As String = Split(csv, ",")
        Console.WriteLine(arr.Length)
        Console.WriteLine(arr(1))
        
        ' Join array into string
        Dim joined = Join(arr, "-")
        Console.WriteLine(joined)
        
        ' Filter array
        Dim words() As String = {"Apple", "Banana", "Cherry", "Apricot"}
        Dim filtered = Filter(words, "Ap")
        Console.WriteLine(filtered.Length)
        Console.WriteLine(filtered(0))
        Console.WriteLine(filtered(1))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["3", "B", "A-B-C", "2", "Apple", "Apricot"]);
}

use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: LINQ OfType(Of T) Mixed Collection Filtering
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_linq_of_type_filter_heterogeneous_array() {
    let src = r#"
Imports System.Collections
Imports System.Linq

Module Program
    Sub Main()
        Dim mixed As Object() = {10, "Hello", 20.5, "World", 30}
        Dim strings = mixed.OfType(Of String)()
        Dim ints = mixed.OfType(Of Integer)()

        Console.WriteLine(String.Join(",", strings))
        Console.WriteLine(String.Join(",", ints))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Hello,World", "10,30"]);
}

#[test]
fn test_vb_linq_of_type_class_hierarchy() {
    let src = r#"
Imports System.Linq

Class Base
End Class

Class ChildA
    Inherits Base
End Class

Class ChildB
    Inherits Base
End Class

Module Program
    Sub Main()
        Dim list As Base() = {New ChildA(), New ChildB(), New ChildA()}
        Dim onlyA = list.OfType(Of ChildA)()
        Console.WriteLine(onlyA.Count())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2"]);
}

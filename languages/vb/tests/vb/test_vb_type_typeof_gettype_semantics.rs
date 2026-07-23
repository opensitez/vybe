use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: GetType Operator vs Object.GetType Method
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_gettype_operator_vs_object_gettype() {
    let src = r#"
Imports System

Class Animal
End Class

Class Dog
    Inherits Animal
End Class

Module Program
    Sub Main()
        Dim d As Animal = New Dog()

        Dim staticType As Type = GetType(Animal)
        Dim runtimeType As Type = d.GetType()

        Console.WriteLine(staticType.Name)
        Console.WriteLine(runtimeType.Name)
        Console.WriteLine(staticType = runtimeType)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Animal", "Dog", "False"]);
}

use super::helpers::run_vb;

#[test]
fn interface_inheritance_basic() {
    let out = run_vb(
        r#"
Interface I1
    Sub M1()
End Interface

Interface I2
    Inherits I1
    Sub M2()
End Interface

Class C
    Implements I2
    
    Public Sub M1() Implements I1.M1
        Console.WriteLine("M1")
    End Sub
    
    Public Sub M2() Implements I2.M2
        Console.WriteLine("M2")
    End Sub
End Class

Module M
    Sub Main()
        Dim obj As I2 = New C()
        obj.M1()
        obj.M2()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["M1", "M2"]);
}

#[test]
fn interface_inheritance_hiding() {
    let out = run_vb(
        r#"
Interface IBase
    Function GetVal() As Integer
End Interface

Interface IDerived
    Inherits IBase
    Shadows Function GetVal() As Integer
End Interface

Class C
    Implements IDerived
    
    Private Function IBase_GetVal() As Integer Implements IBase.GetVal
        Return 1
    End Function
    
    Private Function IDerived_GetVal() As Integer Implements IDerived.GetVal
        Return 2
    End Function
End Class

Module M
    Sub Main()
        Dim obj As New C()
        Dim b As IBase = obj
        Dim d As IDerived = obj
        
        Console.WriteLine(b.GetVal())
        Console.WriteLine(d.GetVal())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["1", "2"]);
}

use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Control flow — With, Using, GoTo, Select Case,
// Exit, Continue, single-line If
// ═══════════════════════════════════════════════════════════

#[test]
fn select_case_basic() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim x As Integer = 2
        Select Case x
            Case 1
                Console.WriteLine("one")
            Case 2
                Console.WriteLine("two")
            Case 3
                Console.WriteLine("three")
            Case Else
                Console.WriteLine("other")
        End Select
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["two"]);
}

#[test]
fn select_case_string() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim color As String = "red"
        Select Case color
            Case "red"
                Console.WriteLine("R")
            Case "green"
                Console.WriteLine("G")
            Case "blue"
                Console.WriteLine("B")
        End Select
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["R"]);
}

#[test]
fn select_case_else() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim x As Integer = 99
        Select Case x
            Case 1
                Console.WriteLine("one")
            Case Else
                Console.WriteLine("default")
        End Select
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["default"]);
}

#[test]
fn if_elseif_chain() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim score As Integer = 75
        If score >= 90 Then
            Console.WriteLine("A")
        ElseIf score >= 80 Then
            Console.WriteLine("B")
        ElseIf score >= 70 Then
            Console.WriteLine("C")
        Else
            Console.WriteLine("F")
        End If
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["C"]);
}

#[test]
fn single_line_if() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim x As Integer = 5
        If x > 3 Then Console.WriteLine("big") Else Console.WriteLine("small")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["big"]);
}

#[test]
fn for_loop_basic() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim sum As Integer = 0
        For i As Integer = 1 To 5
            sum = sum + i
        Next
        Console.WriteLine(sum)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["15"]);
}

#[test]
fn for_loop_step() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim sum As Integer = 0
        For i As Integer = 0 To 10 Step 2
            sum = sum + i
        Next
        Console.WriteLine(sum)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["30"]);
}

#[test]
fn for_loop_negative_step() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        For i As Integer = 5 To 1 Step -1
            Console.WriteLine(i)
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["5", "4", "3", "2", "1"]);
}

#[test]
fn while_loop() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim i As Integer = 0
        While i < 3
            Console.WriteLine(i)
            i = i + 1
        End While
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["0", "1", "2"]);
}

#[test]
fn do_while_loop() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim i As Integer = 0
        Do While i < 3
            Console.WriteLine(i)
            i = i + 1
        Loop
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["0", "1", "2"]);
}

#[test]
fn do_loop_until() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim i As Integer = 0
        Do
            i = i + 1
        Loop Until i >= 3
        Console.WriteLine(i)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn for_each_array() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim arr() As Integer = {10, 20, 30}
        Dim sum As Integer = 0
        For Each x As Integer In arr
            sum = sum + x
        Next
        Console.WriteLine(sum)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["60"]);
}

#[test]
fn exit_for() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        For i As Integer = 1 To 100
            If i > 3 Then Exit For
            Console.WriteLine(i)
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn exit_while() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim i As Integer = 0
        While True
            i = i + 1
            If i = 3 Then Exit While
        End While
        Console.WriteLine(i)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn continue_for() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim sum As Integer = 0
        For i As Integer = 1 To 10
            If i Mod 2 <> 0 Then Continue For
            sum = sum + i
        Next
        Console.WriteLine(sum)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["30"]);
}

#[test]
fn nested_loops() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim count As Integer = 0
        For i As Integer = 1 To 3
            For j As Integer = 1 To 3
                count = count + 1
            Next
        Next
        Console.WriteLine(count)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn with_block() {
    let out = run_vb(
        r#"
Class Person
    Public Name As String
    Public Age As Integer
End Class

Module M
    Sub Main()
        Dim p As New Person()
        With p
            .Name = "Alice"
            .Age = 30
        End With
        Console.WriteLine(p.Name)
        Console.WriteLine(p.Age)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Alice", "30"]);
}

#[test]
fn exit_sub() {
    let out = run_vb(
        r#"
Module M
    Sub Process(x As Integer)
        If x < 0 Then Exit Sub
        Console.WriteLine(x)
    End Sub
    Sub Main()
        Process(5)
        Process(-1)
        Process(10)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["5", "10"]);
}

#[test]
fn exit_function() {
    let out = run_vb(
        r#"
Module M
    Function SafeDiv(a As Integer, b As Integer) As Integer
        If b = 0 Then
            SafeDiv = -1
            Exit Function
        End If
        SafeDiv = a \ b
    End Function
    Sub Main()
        Console.WriteLine(SafeDiv(10, 2))
        Console.WriteLine(SafeDiv(10, 0))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["5", "-1"]);
}

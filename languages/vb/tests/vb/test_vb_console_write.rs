use super::helpers::run_vb;

#[test]
fn writeline_emits_line_per_call() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine("a")
        Console.WriteLine("b")
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["a", "b"]);
}

#[test]
fn write_concatenates_without_newline() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.Write("a")
        Console.Write("b")
        Console.WriteLine()
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["ab"]);
}

#[test]
fn write_then_writeline_appends_to_started_line() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.Write("x")
        Console.WriteLine("y")
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["xy"]);
}

#[test]
fn write_primitive_values_are_rendered_without_separator() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.Write(1)
        Console.Write(2)
        Console.Write(3)
        Console.WriteLine()
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["123"]);
}

#[test]
fn writeline_bool_renders_true_false() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine(True)
        Console.WriteLine(False)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn write_bool_then_text_keeps_same_line() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.Write(False)
        Console.Write("!")
        Console.WriteLine()
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["False!"]);
}

#[test]
fn writeline_no_args_is_a_blank_line() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine()
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec![""]);
}

#[test]
fn writeline_empty_string_keeps_position_between_contents() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine("a")
        Console.WriteLine("")
        Console.WriteLine("b")
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["a", "", "b"]);
}

#[test]
fn write_then_empty_writeline_finalizes_line() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.Write("prefix")
        Console.WriteLine()
        Console.WriteLine("next")
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["prefix", "next"]);
}

#[test]
fn write_and_writeline_mix_over_multiple_lines() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.Write("a")
        Console.Write("b")
        Console.WriteLine("c")
        Console.WriteLine("d")
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["abc", "d"]);
}

#[test]
fn write_single_call_is_captured() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.Write("solo")
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["solo"]);
}

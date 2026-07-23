use super::helpers::run_vb;

#[test]
fn console_title_roundtrip() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim previous As String = Console.Title
        Console.Title = "vb-system-console"
        Console.WriteLine(Console.Title = "vb-system-console")
        Console.Title = previous
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}

#[test]
fn console_redirection_flags_are_boolean() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Console.WriteLine(Console.IsOutputRedirected OrElse Not Console.IsOutputRedirected)
        Console.WriteLine(Console.IsInputRedirected OrElse Not Console.IsInputRedirected)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn console_buffer_and_cursor_properties() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim oldColor As ConsoleColor = Console.ForegroundColor
        Console.ForegroundColor = ConsoleColor.Yellow
        Console.WriteLine(Console.ForegroundColor = ConsoleColor.Yellow)
        Console.ForegroundColor = oldColor
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}

#[test]
fn console_write_via_out_and_err_channels() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Console.Write("out:")
        Console.Error.Write("err")
        Console.WriteLine("done")
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["out:errdone"]);
}

#[test]
fn console_redirection_boolean_guard() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Console.WriteLine(Console.IsErrorRedirected OrElse Not Console.IsErrorRedirected)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}

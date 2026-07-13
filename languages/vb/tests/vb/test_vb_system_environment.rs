use super::helpers::run_vb;

#[test]
fn system_environment_newline() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim nl = Environment.NewLine
        Dim isOk = (nl = vbCrLf OrElse nl = vbLf)
        Console.WriteLine(isOk)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn system_environment_version() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim ver = Environment.Version
        Console.WriteLine(ver IsNot Nothing)
        Console.WriteLine(ver.Major >= 1)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn system_environment_tickcount() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim ticks = Environment.TickCount
        ' Wait 10ms (roughly)
        System.Threading.Thread.Sleep(10)
        Dim ticks2 = Environment.TickCount
        
        Console.WriteLine(ticks2 >= ticks)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["True"]);
}

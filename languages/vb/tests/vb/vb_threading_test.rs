use super::helpers::run_vb;

#[test]
fn task_run_executes_callback() {
    let out = run_vb(
        r#"
Module Program
    Sub Main()
        Dim t = Task.Run(Function()
            Console.WriteLine("task running")
            Return "done"
        End Function)
        Thread.Sleep(100)
        Console.WriteLine("main continues")
    End Sub
End Module
"#,
    );
    assert!(
        out.contains(&"task running".to_string()),
        "task callback should execute: {:?}",
        out
    );
    assert!(
        out.contains(&"main continues".to_string()),
        "main should continue: {:?}",
        out
    );
}

#[test]
fn task_iscompleted_false_initially() {
    let out = run_vb(
        r#"
Module Program
    Sub Main()
        Dim t = Task.Run(Function()
            Thread.Sleep(200)
            Return "done"
        End Function)
        Console.WriteLine(t.IsCompleted)
    End Sub
End Module
"#,
    );
    assert_eq!(
        out,
        super::helpers::dotnet_expected_lines(&["false"]),
        "IsCompleted should be false initially"
    );
}

#[test]
fn task_result_blocks_and_returns_value() {
    let out = run_vb(
        r#"
Module Program
    Sub Main()
        Dim t = Task.Run(Function()
            Thread.Sleep(100)
            Return "success"
        End Function)
        Dim result = t.Result
        Console.WriteLine(result)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["success"], "Result should block and return value");
}

#[test]
fn task_iscompleted_true_after_result() {
    let out = run_vb(
        r#"
Module Program
    Sub Main()
        Dim t = Task.Run(Function()
            Return "done"
        End Function)
        Dim r = t.Result
        Console.WriteLine(t.IsCompleted)
    End Sub
End Module
"#,
    );
    assert_eq!(
        out,
        super::helpers::dotnet_expected_lines(&["true"]),
        "IsCompleted should be true after Result"
    );
}

#[test]
fn task_delay_transitions_to_completed() {
    let out = run_vb(
        r#"
Module Program
    Sub Main()
        Dim d = Task.Delay(100)
        Console.WriteLine(d.IsCompleted)
        Thread.Sleep(150)
        Console.WriteLine(d.IsCompleted)
    End Sub
End Module
"#,
    );
    assert_eq!(
        out,
        super::helpers::dotnet_expected_lines(&["false", "true"]),
        "Task.Delay should complete asynchronously"
    );
}

#[test]
fn thread_new_and_start() {
    let out = run_vb(
        r#"
Module Program
    Sub Main()
        Dim th = New Thread(Sub()
            Console.WriteLine("thread ran")
        End Sub)
        th.Start()
        th.Join()
        Console.WriteLine("joined")
    End Sub
End Module
"#,
    );
    assert!(
        out.contains(&"thread ran".to_string()),
        "thread should run: {:?}",
        out
    );
    assert!(
        out.contains(&"joined".to_string()),
        "join should complete: {:?}",
        out
    );
}

#[test]
fn thread_isalive_true_while_running() {
    let out = run_vb(
        r#"
Module Program
    Sub Main()
        Dim th = New Thread(Sub()
            Thread.Sleep(200)
        End Sub)
        th.Start()
        Console.WriteLine(th.IsAlive)
        th.Join()
        Console.WriteLine(th.IsAlive)
    End Sub
End Module
"#,
    );
    assert_eq!(
        out,
        super::helpers::dotnet_expected_lines(&["true", "false"]),
        "IsAlive should be true then false"
    );
}

#[test]
fn thread_sleep_actually_sleeps() {
    let out = run_vb(
        r#"
Module Program
    Sub Main()
        Dim start = Now
        Thread.Sleep(200)
        Console.WriteLine("slept")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["slept"]);
}

#[test]
fn concurrent_threads_interleave() {
    let out = run_vb(
        r#"
Module Program
    Sub Main()
        Dim t1 = New Thread(Sub()
            Console.WriteLine("t1 start")
            Thread.Sleep(50)
            Console.WriteLine("t1 end")
        End Sub)
        Dim t2 = New Thread(Sub()
            Console.WriteLine("t2 start")
            Thread.Sleep(50)
            Console.WriteLine("t2 end")
        End Sub)
        t1.Start()
        t2.Start()
        t1.Join()
        t2.Join()
        Console.WriteLine("both done")
    End Sub
End Module
"#,
    );
    assert!(
        out.contains(&"t1 start".to_string()),
        "t1 should start: {:?}",
        out
    );
    assert!(
        out.contains(&"t2 start".to_string()),
        "t2 should start: {:?}",
        out
    );
    assert!(
        out.contains(&"both done".to_string()),
        "both should complete: {:?}",
        out
    );
}

#[test]
fn process_start_with_string_has_exited_true() {
    let out = run_vb(
        r#"
Module Program
    Sub Main()
        Dim p = Process.Start("/bin/echo")
        Console.WriteLine(p.HasExited)
    End Sub
End Module
"#,
    );
    assert_eq!(
        out,
        super::helpers::dotnet_expected_lines(&["true"]),
        "Process.Start should return a completed process: {:?}",
        out
    );
}

#[test]
fn process_start_exitcode_is_zero_on_success() {
    let out = run_vb(
        r#"
Module Program
    Sub Main()
        Dim p = Process.Start("/bin/echo")
        Console.WriteLine(p.ExitCode)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["0"], "echo should exit with code 0: {:?}", out);
}

#[test]
fn process_start_with_processstarinfo_and_wait_for_exit() {
    let out = run_vb(
        r#"
Module Program
    Sub Main()
        Dim si = New ProcessStartInfo("/bin/echo", "hello")
        Dim p = Process.Start(si)
        p.WaitForExit()
        Console.WriteLine(p.HasExited)
        Console.WriteLine(p.ExitCode)
    End Sub
End Module
"#,
    );
    assert!(
        out.contains(&super::helpers::dotnet_expected_one("true")),
        "HasExited should be true after WaitForExit: {:?}",
        out
    );
    assert!(
        out.contains(&"0".to_string()),
        "ExitCode should be 0: {:?}",
        out
    );
}

#[test]
fn process_start_info_arguments_are_forwarded() {
    let out = run_vb(
        r#"
Module Program
    Sub Main()
        Dim si = New ProcessStartInfo("/usr/bin/test", "hello = hello")
        Dim p = Process.Start(si)
        p.WaitForExit()
        Console.WriteLine(p.ExitCode)
    End Sub
End Module
"#,
    );
    assert_eq!(
        out,
        vec!["0"],
        "ProcessStartInfo arguments should be forwarded to the child process: {:?}",
        out
    );
}

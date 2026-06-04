use super::helpers::run_csharp;

// ═══════════════════════════════════════════════════════════
// C#: Error handling — try/catch/finally, throw, custom exceptions
// ═══════════════════════════════════════════════════════════

#[test]
fn try_catch_basic() {
    let out = run_csharp(
        r#"
try {
    throw new Exception("oops");
} catch (Exception e) {
    Console.WriteLine(e.Message);
}
"#,
    );
    assert_eq!(out, vec!["oops"]);
}

#[test]
fn try_catch_no_error() {
    let out = run_csharp(
        r#"
try {
    Console.WriteLine("ok");
} catch (Exception e) {
    Console.WriteLine("error");
}
"#,
    );
    assert_eq!(out, vec!["ok"]);
}

#[test]
fn try_finally() {
    let out = run_csharp(
        r#"
try {
    Console.WriteLine("try");
} finally {
    Console.WriteLine("finally");
}
"#,
    );
    assert_eq!(out, vec!["try", "finally"]);
}

#[test]
fn try_catch_finally() {
    let out = run_csharp(
        r#"
try {
    throw new Exception("fail");
} catch (Exception e) {
    Console.WriteLine("caught: " + e.Message);
} finally {
    Console.WriteLine("cleanup");
}
"#,
    );
    assert_eq!(out, vec!["caught: fail", "cleanup"]);
}

#[test]
fn nested_try_catch() {
    let out = run_csharp(
        r#"
try {
    try {
        throw new Exception("inner");
    } catch (Exception e) {
        Console.WriteLine("inner: " + e.Message);
        throw new Exception("rethrown");
    }
} catch (Exception e) {
    Console.WriteLine("outer: " + e.Message);
}
"#,
    );
    assert_eq!(out, vec!["inner: inner", "outer: rethrown"]);
}

#[test]
fn throw_from_method() {
    let out = run_csharp(
        r#"
int Divide(int a, int b) {
    if (b == 0) throw new Exception("Division by zero");
    return a / b;
}
try {
    Console.WriteLine(Divide(10, 2));
    Console.WriteLine(Divide(10, 0));
} catch (Exception e) {
    Console.WriteLine(e.Message);
}
"#,
    );
    assert_eq!(out, vec!["5", "Division by zero"]);
}

#[test]
fn finally_always_runs() {
    let out = run_csharp(
        r#"
try {
    Console.WriteLine("before");
    throw new Exception("err");
} catch (Exception e) {
    Console.WriteLine("caught");
} finally {
    Console.WriteLine("always");
}
"#,
    );
    assert_eq!(out, vec!["before", "caught", "always"]);
}

#[test]
fn exception_message_access() {
    let out = run_csharp(
        r#"
try {
    throw new Exception("test message");
} catch (Exception e) {
    Console.WriteLine(e.Message);
}
"#,
    );
    assert_eq!(out, vec!["test message"]);
}

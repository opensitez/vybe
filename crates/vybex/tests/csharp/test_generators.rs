//! C# generators via `yield return` + WASM stack switching.
//! Any method whose body contains `yield return` (or `yield break`)
//! is marked `chunk.is_generator = true`; calls return a
//! `Continuation`. `foreach` will drive them once iterator-protocol
//! support lands for `foreach`; for now the test exercises the
//! via-program-run path that Python and JS share.

use super::helpers::run_csharp;

#[test]
fn yield_return_emits_continuation() {
    let out = run_csharp(r#"
class Program {
    public static IEnumerable<int> Count() {
        yield return 1;
        yield return 2;
        yield return 3;
    }
    public static void Run() {
        var g = Count();
        Console.WriteLine(g);
    }
}
Program.Run();
"#);
    assert_eq!(out, vec!["[continuation]"]);
}

#[test]
fn yield_return_body_does_not_eagerly_run() {
    let out = run_csharp(r#"
class Program {
    public static IEnumerable<int> Loud() {
        Console.WriteLine("bad: body ran without resume");
        yield return 1;
    }
    public static void Run() {
        var _ = Loud();
        Console.WriteLine("ok");
    }
}
Program.Run();
"#);
    assert_eq!(out, vec!["ok"]);
}

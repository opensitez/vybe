// vybe-test: csharp/csharp_switch_expressions/switch_expression_handles_enum_like_constants
// origin: languages/csharp/tests/csharp/test_csharp_switch_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum State { Idle, Running, Done } var state = State.Done; __Check((state switch { State.Idle => "idle", State.Running => "running", State.Done => "done", _ => "other" }).ToString(), "done");

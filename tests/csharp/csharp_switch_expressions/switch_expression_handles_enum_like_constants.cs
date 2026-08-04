// vybe-test: csharp/csharp_switch_expressions/switch_expression_handles_enum_like_constants
// origin: languages/csharp/tests/csharp/test_csharp_switch_expressions.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

enum State { Idle, Running, Done } var state = State.Done; __P((state switch { State.Idle => "idle", State.Running => "running", State.Done => "done", _ => "other" }).ToString());
__Check("done");

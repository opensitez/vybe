// vybe-test: csharp/csharp_switch_expressions/switch_expression_handles_enum_like_constants
// origin: languages/csharp/tests/csharp/test_csharp_switch_expressions.rs

using static __Harness;

var state = State.Done;
__P((state switch { State.Idle => "idle", State.Running => "running", State.Done => "done", _ => "other" }).ToString());
__Check("done");

enum State { Idle, Running, Done }

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}

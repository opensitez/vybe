// vybe-test: csharp/csharp_using_declarations/using_var_in_switch_case_block_disposes_on_case_exit
// origin: languages/csharp/tests/csharp/test_csharp_using_declarations.rs

using static __Harness;

RunScope();
__Check("case\nsw\nafter");

void RunScope() {
    using var r = new DisposableTracker("case\nsw\nafter");
}

class DisposableTracker : IDisposable {
    private string msg;
    public DisposableTracker(string msg) => this.msg = msg;
    public void Dispose() => __P(msg);
}
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

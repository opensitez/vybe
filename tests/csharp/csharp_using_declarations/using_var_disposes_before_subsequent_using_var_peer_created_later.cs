// vybe-test: csharp/csharp_using_declarations/using_var_disposes_before_subsequent_using_var_peer_created_later
// origin: languages/csharp/tests/csharp/test_csharp_using_declarations.rs

using static __Harness;

RunScope();
__Check("late\nend\nearly");

void RunScope() {
    using var r = new DisposableTracker("late\nend\nearly");
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

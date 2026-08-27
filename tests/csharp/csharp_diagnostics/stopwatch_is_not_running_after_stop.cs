// vybe-test: csharp/csharp_diagnostics/stopwatch_is_not_running_after_stop
// origin: languages/csharp/tests/csharp/test_csharp_diagnostics.rs

using static __Harness;

var sw=System.Diagnostics.Stopwatch.StartNew();
sw.Stop();
__P((sw.IsRunning).ToString());
__Check("False");

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

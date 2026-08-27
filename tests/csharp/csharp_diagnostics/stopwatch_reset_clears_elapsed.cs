// vybe-test: csharp/csharp_diagnostics/stopwatch_reset_clears_elapsed
// origin: languages/csharp/tests/csharp/test_csharp_diagnostics.rs

using static __Harness;

var sw=System.Diagnostics.Stopwatch.StartNew();
System.Threading.Thread.Sleep(5);
sw.Stop();
sw.Reset();
__P((sw.Elapsed==System.TimeSpan.Zero).ToString());
__Check("True");

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

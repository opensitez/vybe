// vybe-test: csharp/csharp_iterators/yield_break_stops_iteration_early
// origin: languages/csharp/tests/csharp/test_csharp_iterators.rs

using static __Harness;

System.Collections.Generic.IEnumerable<int> Gen() {
    yield return 1;
    yield break;
    yield return 2;
}
int count = 0;
foreach(var _ in Gen()) count++;
__P((count).ToString());
__Check("1");

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

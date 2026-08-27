// vybe-test: csharp/csharp_iterators/yield_return_sequence_consumed_by_foreach
// origin: languages/csharp/tests/csharp/test_csharp_iterators.rs

using static __Harness;

System.Collections.Generic.IEnumerable<int> Gen() {
    yield return 1; yield return 2; yield return 3;
}
int sum = 0;
foreach(var n in Gen()) sum += n;
__P((sum).ToString());
__Check("6");

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

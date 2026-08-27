// vybe-test: csharp/csharp_yield_iterators_core/yield_return_from_static_method
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

using static __Harness;

__P((string.Join(",",Seq.Twice(5))).ToString());
__Check("5,10");

class Seq{public static System.Collections.Generic.IEnumerable<int> Twice(int n){yield return n;yield return n*2;}}

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

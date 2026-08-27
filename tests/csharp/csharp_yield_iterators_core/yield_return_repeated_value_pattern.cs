// vybe-test: csharp/csharp_yield_iterators_core/yield_return_repeated_value_pattern
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

using static __Harness;

System.Collections.Generic.IEnumerable<int> Repeat(int v,int n){for(int i=0;i<n;i++)yield return v;}
__P((string.Join(",",Repeat(7,3))).ToString());
__Check("7,7,7");

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

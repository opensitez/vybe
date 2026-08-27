// vybe-test: csharp/csharp_yield_iterators_core/yield_return_break_on_condition_in_foreach_source
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

using static __Harness;

System.Collections.Generic.IEnumerable<int> TakeWhilePositive(int[] a){foreach(var n in a){if(n<0)yield break;yield return n;}}
__P((string.Join(",",TakeWhilePositive(new[]{2,4,-1,8}))).ToString());
__Check("2,4");

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

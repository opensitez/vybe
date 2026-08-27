// vybe-test: csharp/csharp_yield_iterators_core/nested_yield_with_conditional_inner_skip
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

using static __Harness;

System.Collections.Generic.IEnumerable<int> Inner(bool ok){if(ok)yield return 9;}
System.Collections.Generic.IEnumerable<int> Outer(bool ok){foreach(var x in Inner(ok))yield return x;yield return 1;}
__P((string.Join(",",Outer(false))).ToString());
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

// vybe-test: csharp/csharp_yield_iterators_core/yield_return_nested_try_finally_inner_finally_print
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

using static __Harness;

System.Collections.Generic.IEnumerable<int> Gen(){try{try{yield return 1;}finally{__P(("inner").ToString());}}finally{__P(("outer").ToString());}}
foreach(var _ in Gen()){}
__Check("inner\nouter");

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

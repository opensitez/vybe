// vybe-test: csharp/csharp_yield_iterators_core/iterator_finally_runs_after_partial_to_list
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

using static __Harness;

System.Collections.Generic.IEnumerable<int> Gen(){try{yield return 1;yield return 2;yield return 3;}finally{__P(("close").ToString());}}
Gen().Take(2).ToList();
__Check("close");

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

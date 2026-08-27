// vybe-test: csharp/csharp_yield_iterators_core/iterator_finally_prints_even_if_no_yield_reached
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

using static __Harness;

System.Collections.Generic.IEnumerable<int> Gen(bool ok){try{if(!ok)yield break;yield return 1;}finally{__P(("end").ToString());}}
foreach(var _ in Gen(false)){}
__Check("end");

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

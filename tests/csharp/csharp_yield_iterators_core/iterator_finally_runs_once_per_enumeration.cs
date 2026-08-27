// vybe-test: csharp/csharp_yield_iterators_core/iterator_finally_runs_once_per_enumeration
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

using static __Harness;

int fin=0;
System.Collections.Generic.IEnumerable<int> Gen(){try{yield return 1;}finally{fin++;__P((fin).ToString());}}
foreach(var _ in Gen()){}
foreach(var _ in Gen()){}
__Check("1\n2");

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

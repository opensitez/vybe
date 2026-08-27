// vybe-test: csharp/csharp_yield_iterators_core/iterator_finally_not_run_until_started_iteration
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

using static __Harness;

int fin=0;
System.Collections.Generic.IEnumerable<int> Gen(){try{yield return 1;}finally{fin=1;__P((fin).ToString());}}
var seq=Gen();
__P((fin).ToString());
foreach(var _ in seq){}
__Check("0\n1");

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

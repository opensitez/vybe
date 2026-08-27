// vybe-test: csharp/csharp_yield_iterators_core/iterator_disposal_finally_print_once_per_full_run
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

using static __Harness;

int hits=0;
System.Collections.Generic.IEnumerable<int> Gen(){try{yield return 1;}finally{hits++;__P((hits).ToString());}}
foreach(var _ in Gen()){}
__P((hits).ToString());
__Check("1\n1");

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

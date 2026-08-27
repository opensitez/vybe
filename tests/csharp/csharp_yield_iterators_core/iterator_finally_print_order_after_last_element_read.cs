// vybe-test: csharp/csharp_yield_iterators_core/iterator_finally_print_order_after_last_element_read
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

using static __Harness;

System.Collections.Generic.IEnumerable<int> Gen(){try{yield return 10;yield return 20;}finally{__P(("after").ToString());}}
foreach(var n in Gen()) __P((n).ToString());
__Check("10\n20\nafter");

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

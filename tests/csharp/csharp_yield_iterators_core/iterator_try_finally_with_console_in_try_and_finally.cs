// vybe-test: csharp/csharp_yield_iterators_core/iterator_try_finally_with_console_in_try_and_finally
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

using static __Harness;

System.Collections.Generic.IEnumerable<int> Gen(){try{__P(("try").ToString());yield return 5;}finally{__P(("finally").ToString());}}
foreach(var n in Gen()) __P((n).ToString());
__Check("try\n5\nfinally");

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

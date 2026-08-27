// vybe-test: csharp/csharp_yield_iterators_core/nested_iterator_select_many_style_flatten
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

using static __Harness;

System.Collections.Generic.IEnumerable<int> Pair(int n){yield return n;yield return n+1;}
__P((string.Join(",",new[]{1,2}.SelectMany(Pair))).ToString());
__Check("1,2,2,3");

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

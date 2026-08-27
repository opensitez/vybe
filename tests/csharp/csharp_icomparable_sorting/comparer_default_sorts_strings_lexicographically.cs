// vybe-test: csharp/csharp_icomparable_sorting/comparer_default_sorts_strings_lexicographically
// origin: languages/csharp/tests/csharp/test_csharp_icomparable_sorting.rs

using static __Harness;

var list = new System.Collections.Generic.List<string>{"banana","apple","cherry"}
;
list.Sort(System.StringComparer.Ordinal);
__P((list[0]).ToString());
__Check("apple");

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

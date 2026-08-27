// vybe-test: csharp/csharp_string_comparison/string_comparer_ordinal_ignore_case_used_in_sorted_set
// origin: languages/csharp/tests/csharp/test_csharp_string_comparison.rs

using static __Harness;

var set = new System.Collections.Generic.SortedSet<string>(
    System.StringComparer.OrdinalIgnoreCase);
set.Add("Apple");
set.Add("apple");
__P((set.Count).ToString());
__Check("1");

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

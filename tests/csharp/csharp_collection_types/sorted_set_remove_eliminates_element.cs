// vybe-test: csharp/csharp_collection_types/sorted_set_remove_eliminates_element
// origin: languages/csharp/tests/csharp/test_csharp_collection_types.rs

using static __Harness;

var s=new System.Collections.Generic.SortedSet<int>{1,2,3}
;
s.Remove(2);
__P((s.Count).ToString());
__Check("2");

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

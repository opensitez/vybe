// vybe-test: csharp/csharp_generic_collections/sorted_list_maintains_key_order_on_insertion
// origin: languages/csharp/tests/csharp/test_csharp_generic_collections.rs

using static __Harness;

var sl = new System.Collections.Generic.SortedList<int,string>();
sl.Add(3,"c");
sl.Add(1,"a");
sl.Add(2,"b");
__P((sl.Keys[0]).ToString());
__P((sl.Values[0]).ToString());
__Check("1\na");

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

// vybe-test: csharp/csharp_generic_collections/sorted_list_index_of_key_finds_insertion_position
// origin: languages/csharp/tests/csharp/test_csharp_generic_collections.rs

using static __Harness;

var sl = new System.Collections.Generic.SortedList<string,int>{{"a",1},{"b",2},{"c",3}}
;
__P((sl.IndexOfKey("b")).ToString());
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

// vybe-test: csharp/csharp_list_operations/insert_places_element_at_specified_index
// origin: languages/csharp/tests/csharp/test_csharp_list_operations.rs

using static __Harness;

var list = new System.Collections.Generic.List<int>{1,3}
;
list.Insert(1, 2);
__P((list[1]).ToString());
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

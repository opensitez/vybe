// vybe-test: csharp/csharp_list_operations/remove_at_deletes_element_by_index
// origin: languages/csharp/tests/csharp/test_csharp_list_operations.rs

using static __Harness;

var list = new System.Collections.Generic.List<string>{"a","b","c"}
;
list.RemoveAt(0);
__P((list[0]).ToString());
__Check("b");

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

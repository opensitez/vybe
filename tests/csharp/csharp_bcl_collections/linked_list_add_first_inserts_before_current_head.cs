// vybe-test: csharp/csharp_bcl_collections/linked_list_add_first_inserts_before_current_head
// origin: languages/csharp/tests/csharp/test_csharp_bcl_collections.rs

using static __Harness;

var list = new System.Collections.Generic.LinkedList<int>();
list.AddLast(2);
list.AddFirst(1);
__P((list.First.Value).ToString());
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

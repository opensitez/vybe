// vybe-test: csharp/csharp_target_typed_new_delegate/target_new_linked_list_inferred_element_type
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

using static __Harness;

System.Collections.Generic.LinkedList<int> list = new();
list.AddLast(10);
list.AddLast(20);
__P((list.First.Value).ToString());
__P((list.Last.Value).ToString());
__Check("10\n20");

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

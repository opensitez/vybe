// vybe-test: csharp/csharp_collection_types/stack_push_pop_maintains_lifo
// origin: languages/csharp/tests/csharp/test_csharp_collection_types.rs

using static __Harness;

var s=new System.Collections.Generic.Stack<string>();
s.Push("a");
s.Push("b");
__P((s.Pop()).ToString());
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

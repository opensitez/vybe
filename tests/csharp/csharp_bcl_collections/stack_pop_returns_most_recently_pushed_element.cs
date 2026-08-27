// vybe-test: csharp/csharp_bcl_collections/stack_pop_returns_most_recently_pushed_element
// origin: languages/csharp/tests/csharp/test_csharp_bcl_collections.rs

using static __Harness;

var stack = new System.Collections.Generic.Stack<int>();
stack.Push(1);
stack.Push(2);
__P((stack.Pop()).ToString());
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

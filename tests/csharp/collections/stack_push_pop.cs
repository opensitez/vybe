// vybe-test: csharp/collections/stack_push_pop
// origin: languages/csharp/tests/csharp/test_collections.rs

using static __Harness;

var s = new Stack<int>();
s.Push(1);
s.Push(2);
s.Push(3);
__P((s.Pop()).ToString());
__P((s.Pop()).ToString());
__Check("3\n2");

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

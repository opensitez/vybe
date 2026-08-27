// vybe-test: csharp/csharp_generics/generic_stack_queue
// origin: languages/csharp/tests/csharp/test_csharp_generics.rs

using static __Harness;

var stack = new Stack<string>();
stack.Push("first");
stack.Push("second");
__P((stack.Pop()).ToString());
var queue = new Queue<int>();
queue.Enqueue(1);
queue.Enqueue(2);
__P((queue.Dequeue()).ToString());
__Check("second\n1");

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

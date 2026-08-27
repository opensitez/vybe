// vybe-test: csharp/csharp_queue_stack_linkedlist/queue_string_elements_maintain_insertion_order
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_linkedlist.rs

using static __Harness;
using System.Collections.Generic;

var q = new Queue<string>();
q.Enqueue("a");
q.Enqueue("b");
__P((q.Dequeue()).ToString());
__P((q.Dequeue()).ToString());
__Check("a\nb");

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

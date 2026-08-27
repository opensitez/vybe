// vybe-test: csharp/csharp_queue_stack_linkedlist/queue_dequeue_twice_drains_in_fifo_order
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_linkedlist.rs

using static __Harness;
using System.Collections.Generic;

var q = new Queue<int>();
q.Enqueue(10);
q.Enqueue(20);
q.Enqueue(30);
__P((q.Dequeue()).ToString());
__P((q.Dequeue()).ToString());
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

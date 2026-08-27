// vybe-test: csharp/csharp_collection_initializer_syntax/queue_initializer_enqueues_elements_for_fifo_order
// origin: languages/csharp/tests/csharp/test_csharp_collection_initializer_syntax.rs

using static __Harness;
using System.Collections.Generic;

var queue = new Queue<int>();
queue.Enqueue(1);
queue.Enqueue(2);
__P((queue.Dequeue()).ToString());
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

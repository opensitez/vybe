// vybe-test: csharp/csharp_queue_stack_linkedlist/queue_to_array_preserves_fifo_order
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_linkedlist.rs

using static __Harness;
using System.Collections.Generic;

var q = new Queue<int>();
q.Enqueue(3);
q.Enqueue(1);
q.Enqueue(2);
var arr = q.ToArray();
__P((arr[0]).ToString());
__P((arr[2]).ToString());
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

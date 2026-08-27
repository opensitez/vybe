// vybe-test: csharp/csharp_queue_stack_linkedlist/linkedlist_add_first_on_empty_becomes_sole_node
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_linkedlist.rs

using static __Harness;
using System.Collections.Generic;

var ll = new LinkedList<int>();
ll.AddFirst(7);
__P((ll.First.Value).ToString());
__P((ll.Last.Value).ToString());
__Check("7\n7");

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

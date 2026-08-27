// vybe-test: csharp/csharp_queue_stack_linkedlist/linkedlist_foreach_walks_forward_order
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_linkedlist.rs

using static __Harness;
using System.Collections.Generic;

var ll = new LinkedList<int>();
ll.AddLast(1);
ll.AddLast(2);
ll.AddLast(3);
foreach (var x in ll) __P((x).ToString());
__Check("1\n2\n3");

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

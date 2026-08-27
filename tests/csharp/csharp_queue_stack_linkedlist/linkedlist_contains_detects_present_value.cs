// vybe-test: csharp/csharp_queue_stack_linkedlist/linkedlist_contains_detects_present_value
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_linkedlist.rs

using static __Harness;
using System.Collections.Generic;

var ll = new LinkedList<int>();
ll.AddLast(5);
ll.AddLast(6);
__P((ll.Contains(6)).ToString());
__Check("True");

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

// vybe-test: csharp/csharp_queue_stack_linkedlist/stack_to_array_preserves_lifo_top_at_end
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_linkedlist.rs

using static __Harness;
using System.Collections.Generic;

var s = new Stack<int>();
s.Push(1);
s.Push(2);
s.Push(3);
var arr = s.ToArray();
__P((arr[0]).ToString());
__P((arr[2]).ToString());
__Check("3\n1");

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

// vybe-test: csharp/csharp_collections/queue_operations
// origin: languages/csharp/tests/csharp/test_csharp_collections.rs

using static __Harness;
using System.Collections.Generic;

var q = new Queue<string>();
q.Enqueue("first");
q.Enqueue("second");
q.Enqueue("third");
__P((q.Count).ToString());
__P((q.Dequeue()).ToString());
__P((q.Peek()).ToString());
__Check("3\nfirst\nsecond");

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

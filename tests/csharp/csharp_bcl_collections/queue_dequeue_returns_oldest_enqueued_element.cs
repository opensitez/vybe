// vybe-test: csharp/csharp_bcl_collections/queue_dequeue_returns_oldest_enqueued_element
// origin: languages/csharp/tests/csharp/test_csharp_bcl_collections.rs

using static __Harness;

var queue = new System.Collections.Generic.Queue<int>();
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

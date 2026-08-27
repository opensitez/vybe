// vybe-test: csharp/csharp_generic_collections/priority_queue_dequeue_returns_lowest_priority_first
// origin: languages/csharp/tests/csharp/test_csharp_generic_collections.rs

using static __Harness;

var pq = new System.Collections.Generic.PriorityQueue<string,int>();
pq.Enqueue("low", 10);
pq.Enqueue("high", 1);
pq.Enqueue("mid", 5);
__P((pq.Dequeue()).ToString());
__Check("high");

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

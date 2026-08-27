// vybe-test: csharp/csharp_collection_types/queue_enqueue_dequeue_maintains_fifo
// origin: languages/csharp/tests/csharp/test_csharp_collection_types.rs

using static __Harness;

var q=new System.Collections.Generic.Queue<string>();
q.Enqueue("first");
q.Enqueue("second");
__P((q.Dequeue()).ToString());
__Check("first");

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

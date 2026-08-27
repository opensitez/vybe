// vybe-test: csharp/collections/queue_enqueue_dequeue
// origin: languages/csharp/tests/csharp/test_collections.rs

using static __Harness;

var q = new Queue<string>();
q.Enqueue("first");
q.Enqueue("second");
__P((q.Dequeue()).ToString());
__P((q.Dequeue()).ToString());
__Check("first\nsecond");

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

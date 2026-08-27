// vybe-test: csharp/collections_advanced/queue_enqueue_dequeue
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

using static __Harness;

var q = new Queue<string>();
q.Enqueue("first");
q.Enqueue("second");
q.Enqueue("third");
__P((q.Dequeue()).ToString());
__P((q.Dequeue()).ToString());
__P((q.Count).ToString());
__Check("first\nsecond\n1");

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

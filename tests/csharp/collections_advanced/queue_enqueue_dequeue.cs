// vybe-test: csharp/collections_advanced/queue_enqueue_dequeue
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

var q = new Queue<string>();
q.Enqueue("first");
q.Enqueue("second");
q.Enqueue("third");
__P((q.Dequeue()).ToString());
__P((q.Dequeue()).ToString());
__P((q.Count).ToString());
__Check("first\nsecond\n1");

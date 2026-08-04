// vybe-test: csharp/csharp_collection_types/queue_enqueue_dequeue_maintains_fifo
// origin: languages/csharp/tests/csharp/test_csharp_collection_types.rs

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

var q=new System.Collections.Generic.Queue<string>();
q.Enqueue("first"); q.Enqueue("second");
__P((q.Dequeue()).ToString());
__Check("first");

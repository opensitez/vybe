// vybe-test: csharp/csharp_map_set_collections/queue_enqueue_and_dequeue_follow_fifo_order
// origin: languages/csharp/tests/csharp/test_csharp_map_set_collections.rs

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

using System.Collections.Generic; var queue = new Queue<int>(); queue.Enqueue(1); queue.Enqueue(2); __P((queue.Dequeue()).ToString()); __P((queue.Dequeue()).ToString());
__Check("1\n2");

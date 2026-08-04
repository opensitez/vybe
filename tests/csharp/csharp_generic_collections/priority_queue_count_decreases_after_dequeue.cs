// vybe-test: csharp/csharp_generic_collections/priority_queue_count_decreases_after_dequeue
// origin: languages/csharp/tests/csharp/test_csharp_generic_collections.rs

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

var pq = new System.Collections.Generic.PriorityQueue<int,int>();
pq.Enqueue(1,1); pq.Enqueue(2,2);
pq.Dequeue();
__P((pq.Count).ToString());
__Check("1");

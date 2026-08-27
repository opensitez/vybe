// vybe-test: csharp/csharp_collections_priority_queue_min_heap/priority_queue_min_heap_case_17

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

var pq = new System.Collections.Generic.PriorityQueue<string, int>();
pq.Enqueue("Low_17", 100);
pq.Enqueue("High_17", 1);
string first = pq.Dequeue();
__P(first);
__Check("High_17");

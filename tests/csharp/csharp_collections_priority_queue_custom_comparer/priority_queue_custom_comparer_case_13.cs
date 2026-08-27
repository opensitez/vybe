// vybe-test: csharp/csharp_collections_priority_queue_custom_comparer/priority_queue_custom_comparer_case_13

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

var pq = new System.Collections.Generic.PriorityQueue<string, int>(System.Collections.Generic.Comparer<int>.Create((x, y) => y.CompareTo(x)));
pq.Enqueue("Small_13", 1);
pq.Enqueue("Large_13", 100);
string first = pq.Dequeue();
__P(first);
__Check("Large_13");

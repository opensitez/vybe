// vybe-test: csharp/csharp_concurrent_queue_fifo_lifecycle/concurrent_queue_fifo_case_1

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

var cq = new System.Collections.Concurrent.ConcurrentQueue<int>();
cq.Enqueue(1);
cq.Enqueue(11);
bool d1 = cq.TryDequeue(out int v1);
bool d2 = cq.TryDequeue(out int v2);
__P(d1.ToString());
__P(v1.ToString());
__P(d2.ToString());
__P(v2.ToString());
__Check("True\n1\nTrue\n11");

// vybe-test: csharp/csharp_generic_collections/priority_queue_dequeue_returns_lowest_priority_first
// origin: languages/csharp/tests/csharp/test_csharp_generic_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var pq = new System.Collections.Generic.PriorityQueue<string,int>();
pq.Enqueue("low", 10);
pq.Enqueue("high", 1);
pq.Enqueue("mid", 5);
__Check((pq.Dequeue()).ToString(), "high");

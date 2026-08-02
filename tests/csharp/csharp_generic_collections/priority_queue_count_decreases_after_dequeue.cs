// vybe-test: csharp/csharp_generic_collections/priority_queue_count_decreases_after_dequeue
// origin: languages/csharp/tests/csharp/test_csharp_generic_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var pq = new System.Collections.Generic.PriorityQueue<int,int>();
pq.Enqueue(1,1); pq.Enqueue(2,2);
pq.Dequeue();
__Check((pq.Count).ToString(), "1");

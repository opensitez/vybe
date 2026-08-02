// vybe-test: csharp/csharp_bcl_collections/queue_dequeue_returns_oldest_enqueued_element
// origin: languages/csharp/tests/csharp/test_csharp_bcl_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var queue = new System.Collections.Generic.Queue<int>();
queue.Enqueue(1);
queue.Enqueue(2);
__Check((queue.Dequeue()).ToString(), "1");

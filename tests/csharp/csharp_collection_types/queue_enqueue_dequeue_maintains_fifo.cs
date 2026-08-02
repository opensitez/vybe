// vybe-test: csharp/csharp_collection_types/queue_enqueue_dequeue_maintains_fifo
// origin: languages/csharp/tests/csharp/test_csharp_collection_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var q=new System.Collections.Generic.Queue<string>();
q.Enqueue("first"); q.Enqueue("second");
__Check((q.Dequeue()).ToString(), "first");

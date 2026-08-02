// vybe-test: csharp/collections/queue_enqueue_dequeue
// origin: languages/csharp/tests/csharp/test_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var q = new Queue<string>();
        q.Enqueue("first");
        q.Enqueue("second");
        __Check((q.Dequeue()).ToString(), "first");
        __Check((q.Dequeue()).ToString(), "second");

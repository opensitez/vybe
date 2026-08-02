// vybe-test: csharp/collections_advanced/queue_enqueue_dequeue
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var q = new Queue<string>();
q.Enqueue("first");
q.Enqueue("second");
q.Enqueue("third");
__Check((q.Dequeue()).ToString(), "first");
__Check((q.Dequeue()).ToString(), "second");
__Check((q.Count).ToString(), "1");

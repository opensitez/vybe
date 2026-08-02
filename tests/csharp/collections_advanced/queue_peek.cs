// vybe-test: csharp/collections_advanced/queue_peek
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var q = new Queue<int>();
q.Enqueue(10);
q.Enqueue(20);
__Check((q.Peek()).ToString(), "10");
__Check((q.Count).ToString(), "2");

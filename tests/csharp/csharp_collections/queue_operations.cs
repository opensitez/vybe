// vybe-test: csharp/csharp_collections/queue_operations
// origin: languages/csharp/tests/csharp/test_csharp_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic;
var q = new Queue<string>();
q.Enqueue("first");
q.Enqueue("second");
q.Enqueue("third");
__Check((q.Count).ToString(), "3");
__Check((q.Dequeue()).ToString(), "first");
__Check((q.Peek()).ToString(), "second");

// vybe-test: csharp/csharp_queue_stack_linkedlist/queue_dequeue_returns_oldest_element
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_linkedlist.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var q = new Queue<string>(); q.Enqueue("first"); q.Enqueue("second"); __Check((q.Dequeue()).ToString(), "first");

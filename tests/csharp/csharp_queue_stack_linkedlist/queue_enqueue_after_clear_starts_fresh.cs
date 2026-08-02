// vybe-test: csharp/csharp_queue_stack_linkedlist/queue_enqueue_after_clear_starts_fresh
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_linkedlist.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var q = new Queue<int>(); q.Enqueue(1); q.Clear(); q.Enqueue(9); __Check((q.Dequeue()).ToString(), "9");

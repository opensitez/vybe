// vybe-test: csharp/csharp_queue_stack_linkedlist/queue_dequeue_twice_drains_in_fifo_order
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_linkedlist.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var q = new Queue<int>(); q.Enqueue(10); q.Enqueue(20); q.Enqueue(30); __Check((q.Dequeue()).ToString(), "10"); __Check((q.Dequeue()).ToString(), "20");

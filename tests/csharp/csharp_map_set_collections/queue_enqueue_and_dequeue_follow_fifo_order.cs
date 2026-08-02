// vybe-test: csharp/csharp_map_set_collections/queue_enqueue_and_dequeue_follow_fifo_order
// origin: languages/csharp/tests/csharp/test_csharp_map_set_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var queue = new Queue<int>(); queue.Enqueue(1); queue.Enqueue(2); __Check((queue.Dequeue()).ToString(), "1"); __Check((queue.Dequeue()).ToString(), "2");

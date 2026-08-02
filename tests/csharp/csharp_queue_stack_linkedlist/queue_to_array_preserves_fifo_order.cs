// vybe-test: csharp/csharp_queue_stack_linkedlist/queue_to_array_preserves_fifo_order
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_linkedlist.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var q = new Queue<int>(); q.Enqueue(3); q.Enqueue(1); q.Enqueue(2); var arr = q.ToArray(); __Check((arr[0]).ToString(), "3"); __Check((arr[2]).ToString(), "2");

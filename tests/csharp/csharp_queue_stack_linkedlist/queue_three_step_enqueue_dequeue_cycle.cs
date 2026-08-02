// vybe-test: csharp/csharp_queue_stack_linkedlist/queue_three_step_enqueue_dequeue_cycle
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_linkedlist.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var q = new Queue<int>(); q.Enqueue(1); q.Enqueue(2); q.Enqueue(3); q.Dequeue(); __Check((q.Peek()).ToString(), "2"); __Check((q.Count).ToString(), "2");

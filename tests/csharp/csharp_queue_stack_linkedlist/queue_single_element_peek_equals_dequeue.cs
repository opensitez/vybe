// vybe-test: csharp/csharp_queue_stack_linkedlist/queue_single_element_peek_equals_dequeue
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_linkedlist.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var q = new Queue<int>(); q.Enqueue(42); __Check((q.Peek()).ToString(), "42"); __Check((q.Dequeue()).ToString(), "42");

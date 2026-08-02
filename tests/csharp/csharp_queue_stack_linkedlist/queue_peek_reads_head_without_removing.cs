// vybe-test: csharp/csharp_queue_stack_linkedlist/queue_peek_reads_head_without_removing
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_linkedlist.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var q = new Queue<int>(); q.Enqueue(5); q.Enqueue(6); __Check((q.Peek()).ToString(), "5"); __Check((q.Count).ToString(), "2");

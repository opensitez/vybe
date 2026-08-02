// vybe-test: csharp/csharp_queue_stack_linkedlist/queue_contains_finds_enqueued_value
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_linkedlist.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var q = new Queue<int>(); q.Enqueue(7); q.Enqueue(8); __Check((q.Contains(8)).ToString(), "True");

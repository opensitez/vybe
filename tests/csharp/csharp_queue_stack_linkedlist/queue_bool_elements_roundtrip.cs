// vybe-test: csharp/csharp_queue_stack_linkedlist/queue_bool_elements_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_linkedlist.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var q = new Queue<bool>(); q.Enqueue(true); q.Enqueue(false); __Check((q.Dequeue()).ToString(), "True"); __Check((q.Dequeue()).ToString(), "False");

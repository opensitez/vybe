// vybe-test: csharp/csharp_queue_stack_linkedlist/linkedlist_add_first_inserts_at_head
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_linkedlist.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var ll = new LinkedList<int>(); ll.AddLast(2); ll.AddFirst(1); __Check((ll.First.Value).ToString(), "1");

// vybe-test: csharp/csharp_queue_stack_linkedlist/linkedlist_add_last_appends_at_tail
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_linkedlist.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var ll = new LinkedList<int>(); ll.AddFirst(1); ll.AddLast(3); __Check((ll.Last.Value).ToString(), "3");

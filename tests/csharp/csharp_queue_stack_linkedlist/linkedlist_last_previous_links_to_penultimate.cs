// vybe-test: csharp/csharp_queue_stack_linkedlist/linkedlist_last_previous_links_to_penultimate
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_linkedlist.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var ll = new LinkedList<int>(); ll.AddLast(10); ll.AddLast(20); ll.AddLast(30); __Check((ll.Last.Previous.Value).ToString(), "20");

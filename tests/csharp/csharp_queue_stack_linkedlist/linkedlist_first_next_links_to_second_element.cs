// vybe-test: csharp/csharp_queue_stack_linkedlist/linkedlist_first_next_links_to_second_element
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_linkedlist.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var ll = new LinkedList<int>(); ll.AddLast(10); ll.AddLast(20); __Check((ll.First.Next.Value).ToString(), "20");

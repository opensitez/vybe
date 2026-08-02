// vybe-test: csharp/csharp_queue_stack_linkedlist/linkedlist_find_returns_first_of_duplicates
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_linkedlist.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var ll = new LinkedList<int>(); ll.AddLast(1); ll.AddLast(2); ll.AddLast(2); var node = ll.Find(2); __Check((node.Value).ToString(), "2"); __Check((node.Next.Value).ToString(), "2");

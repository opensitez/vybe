// vybe-test: csharp/csharp_queue_stack_linkedlist/linkedlist_add_after_inserts_between_nodes
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_linkedlist.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var ll = new LinkedList<int>(); var n = ll.AddFirst(1); ll.AddAfter(n, 3); ll.AddAfter(n, 2); __Check((n.Next.Value).ToString(), "2");

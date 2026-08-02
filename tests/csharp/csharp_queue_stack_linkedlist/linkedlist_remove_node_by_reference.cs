// vybe-test: csharp/csharp_queue_stack_linkedlist/linkedlist_remove_node_by_reference
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_linkedlist.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var ll = new LinkedList<int>(); var mid = ll.AddLast(2); ll.AddFirst(1); ll.AddLast(3); ll.Remove(mid); __Check((ll.Count).ToString(), "2");

// vybe-test: csharp/csharp_queue_stack_linkedlist/linkedlist_add_before_inserts_predecessor
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_linkedlist.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var ll = new LinkedList<int>(); var tail = ll.AddLast(3); ll.AddBefore(tail, 2); ll.AddBefore(tail, 1); __Check((ll.First.Value).ToString(), "1");

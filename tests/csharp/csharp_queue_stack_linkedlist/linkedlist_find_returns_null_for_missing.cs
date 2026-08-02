// vybe-test: csharp/csharp_queue_stack_linkedlist/linkedlist_find_returns_null_for_missing
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_linkedlist.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var ll = new LinkedList<int>(); ll.AddLast(1); __Check((ll.Find(9) == null).ToString(), "True");

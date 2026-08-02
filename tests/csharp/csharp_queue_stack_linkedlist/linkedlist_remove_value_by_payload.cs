// vybe-test: csharp/csharp_queue_stack_linkedlist/linkedlist_remove_value_by_payload
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_linkedlist.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var ll = new LinkedList<int>(); ll.AddLast(1); ll.AddLast(2); ll.AddLast(3); ll.Remove(2); __Check((ll.Contains(2)).ToString(), "False");

// vybe-test: csharp/csharp_queue_stack_linkedlist/linkedlist_add_last_on_empty_sets_head_and_tail
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_linkedlist.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var ll = new LinkedList<string>(); ll.AddLast("solo"); __Check((ll.First.Value).ToString(), "solo"); __Check((ll.Last.Value).ToString(), "solo");

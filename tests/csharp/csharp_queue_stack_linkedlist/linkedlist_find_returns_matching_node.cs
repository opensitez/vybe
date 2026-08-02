// vybe-test: csharp/csharp_queue_stack_linkedlist/linkedlist_find_returns_matching_node
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_linkedlist.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var ll = new LinkedList<string>(); ll.AddLast("a"); ll.AddLast("target"); var node = ll.Find("target"); __Check((node.Value).ToString(), "target");

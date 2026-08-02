// vybe-test: csharp/csharp_queue_stack_linkedlist/linkedlist_foreach_walks_forward_order
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_linkedlist.rs

using System.Collections.Generic; var ll = new LinkedList<int>(); ll.AddLast(1); ll.AddLast(2); ll.AddLast(3); foreach (var x in ll) Console.WriteLine(x);

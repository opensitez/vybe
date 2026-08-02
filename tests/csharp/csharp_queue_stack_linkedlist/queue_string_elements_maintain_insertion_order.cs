// vybe-test: csharp/csharp_queue_stack_linkedlist/queue_string_elements_maintain_insertion_order
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_linkedlist.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var q = new Queue<string>(); q.Enqueue("a"); q.Enqueue("b"); __Check((q.Dequeue()).ToString(), "a"); __Check((q.Dequeue()).ToString(), "b");

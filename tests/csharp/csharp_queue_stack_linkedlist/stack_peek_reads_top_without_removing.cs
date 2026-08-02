// vybe-test: csharp/csharp_queue_stack_linkedlist/stack_peek_reads_top_without_removing
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_linkedlist.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var s = new Stack<int>(); s.Push(4); s.Push(5); __Check((s.Peek()).ToString(), "5"); __Check((s.Count).ToString(), "2");

// vybe-test: csharp/csharp_queue_stack_linkedlist/stack_push_after_clear_restarts_sequence
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_linkedlist.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var s = new Stack<int>(); s.Push(1); s.Clear(); s.Push(8); __Check((s.Pop()).ToString(), "8");

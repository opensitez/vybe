// vybe-test: csharp/csharp_queue_stack_linkedlist/stack_peek_after_pop_shows_new_top
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_linkedlist.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var s = new Stack<int>(); s.Push(1); s.Push(2); s.Pop(); __Check((s.Peek()).ToString(), "1");

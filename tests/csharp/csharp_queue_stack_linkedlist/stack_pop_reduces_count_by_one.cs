// vybe-test: csharp/csharp_queue_stack_linkedlist/stack_pop_reduces_count_by_one
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_linkedlist.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var s = new Stack<int>(); s.Push(1); s.Push(2); s.Pop(); __Check((s.Count).ToString(), "1");

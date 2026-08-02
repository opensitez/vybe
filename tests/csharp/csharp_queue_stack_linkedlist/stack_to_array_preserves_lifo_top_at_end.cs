// vybe-test: csharp/csharp_queue_stack_linkedlist/stack_to_array_preserves_lifo_top_at_end
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_linkedlist.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var s = new Stack<int>(); s.Push(1); s.Push(2); s.Push(3); var arr = s.ToArray(); __Check((arr[0]).ToString(), "3"); __Check((arr[2]).ToString(), "1");

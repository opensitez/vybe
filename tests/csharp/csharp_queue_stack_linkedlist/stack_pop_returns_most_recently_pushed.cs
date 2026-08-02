// vybe-test: csharp/csharp_queue_stack_linkedlist/stack_pop_returns_most_recently_pushed
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_linkedlist.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var s = new Stack<string>(); s.Push("bottom"); s.Push("top"); __Check((s.Pop()).ToString(), "top");

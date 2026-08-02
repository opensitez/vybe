// vybe-test: csharp/csharp_queue_stack_linkedlist/stack_contains_finds_pushed_value
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_linkedlist.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var s = new Stack<int>(); s.Push(11); s.Push(22); __Check((s.Contains(22)).ToString(), "True");

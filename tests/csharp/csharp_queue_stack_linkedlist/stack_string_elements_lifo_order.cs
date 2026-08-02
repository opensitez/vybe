// vybe-test: csharp/csharp_queue_stack_linkedlist/stack_string_elements_lifo_order
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_linkedlist.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var s = new Stack<string>(); s.Push("x"); s.Push("y"); __Check((s.Pop()).ToString(), "y"); __Check((s.Pop()).ToString(), "x");

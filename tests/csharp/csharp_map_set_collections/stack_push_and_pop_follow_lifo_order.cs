// vybe-test: csharp/csharp_map_set_collections/stack_push_and_pop_follow_lifo_order
// origin: languages/csharp/tests/csharp/test_csharp_map_set_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var stack = new Stack<int>(); stack.Push(1); stack.Push(2); __Check((stack.Pop()).ToString(), "2"); __Check((stack.Pop()).ToString(), "1");

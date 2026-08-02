// vybe-test: csharp/csharp_collection_initializer_syntax/stack_initializer_pushes_elements_for_lifo_order
// origin: languages/csharp/tests/csharp/test_csharp_collection_initializer_syntax.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic;
var stack = new Stack<int>();
stack.Push(1);
stack.Push(2);
__Check((stack.Pop()).ToString(), "2");

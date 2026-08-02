// vybe-test: csharp/csharp_target_typed_new_delegate/target_new_stack_inferred_element_type
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Collections.Generic.Stack<int> s = new();
s.Push(1); s.Push(2);
__Check((s.Pop()).ToString(), "2"); __Check((s.Pop()).ToString(), "1");

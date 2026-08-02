// vybe-test: csharp/csharp_target_typed_new_delegate/target_new_stack_string_inferred
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Collections.Generic.Stack<string> s = new();
s.Push("base"); s.Push("top");
__Check((s.Peek()).ToString(), "top");

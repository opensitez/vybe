// vybe-test: csharp/csharp_collection_types/stack_push_pop_maintains_lifo
// origin: languages/csharp/tests/csharp/test_csharp_collection_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var s=new System.Collections.Generic.Stack<string>();
s.Push("a"); s.Push("b");
__Check((s.Pop()).ToString(), "b");

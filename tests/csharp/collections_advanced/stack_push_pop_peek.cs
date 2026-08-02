// vybe-test: csharp/collections_advanced/stack_push_pop_peek
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var s = new Stack<int>();
s.Push(1);
s.Push(2);
s.Push(3);
__Check((s.Peek()).ToString(), "3");
__Check((s.Pop()).ToString(), "3");
__Check((s.Pop()).ToString(), "2");
__Check((s.Count).ToString(), "1");

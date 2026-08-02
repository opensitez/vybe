// vybe-test: csharp/csharp_bcl_collections/stack_pop_returns_most_recently_pushed_element
// origin: languages/csharp/tests/csharp/test_csharp_bcl_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var stack = new System.Collections.Generic.Stack<int>();
stack.Push(1);
stack.Push(2);
__Check((stack.Pop()).ToString(), "2");

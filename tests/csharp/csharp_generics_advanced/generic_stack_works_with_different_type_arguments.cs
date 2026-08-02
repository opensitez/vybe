// vybe-test: csharp/csharp_generics_advanced/generic_stack_works_with_different_type_arguments
// origin: languages/csharp/tests/csharp/test_csharp_generics_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Stack<T> {
    System.Collections.Generic.List<T> items = new();
    public void Push(T v) => items.Add(v);
    public T Pop() { var v = items[items.Count-1]; items.RemoveAt(items.Count-1); return v; }
}
var s = new Stack<string>();
s.Push("a"); s.Push("b");
__Check((s.Pop()).ToString(), "b");

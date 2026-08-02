// vybe-test: csharp/interfaces_generics/generic_stack_implementation
// origin: languages/csharp/tests/csharp/test_interfaces_generics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class MyStack<T> {
    private List<T> items = new List<T>();
    public void Push(T item) { items.Add(item); }
    public T Pop() {
        T item = items[items.Count - 1];
        items.RemoveAt(items.Count - 1);
        return item;
    }
    public int Count { get { return items.Count; } }
}
var s = new MyStack<int>();
s.Push(10);
s.Push(20);
s.Push(30);
__Check((s.Pop()).ToString(), "30");
__Check((s.Pop()).ToString(), "20");
__Check((s.Count).ToString(), "1");

// vybe-test: csharp/interfaces_generics/generic_stack_implementation
// origin: languages/csharp/tests/csharp/test_interfaces_generics.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
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
__P((s.Pop()).ToString());
__P((s.Pop()).ToString());
__P((s.Count).ToString());
__Check("30\n20\n1");

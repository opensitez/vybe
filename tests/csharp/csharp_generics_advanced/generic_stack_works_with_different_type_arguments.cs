// vybe-test: csharp/csharp_generics_advanced/generic_stack_works_with_different_type_arguments
// origin: languages/csharp/tests/csharp/test_csharp_generics_advanced.rs

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

class Stack<T> {
    System.Collections.Generic.List<T> items = new();
    public void Push(T v) => items.Add(v);
    public T Pop() { var v = items[items.Count-1]; items.RemoveAt(items.Count-1); return v; }
}
var s = new Stack<string>();
s.Push("a"); s.Push("b");
__P((s.Pop()).ToString());
__Check("b");

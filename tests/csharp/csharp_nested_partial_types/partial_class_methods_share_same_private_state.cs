// vybe-test: csharp/csharp_nested_partial_types/partial_class_methods_share_same_private_state
// origin: languages/csharp/tests/csharp/test_csharp_nested_partial_types.rs

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

partial class Counter {
    int value;
}
partial class Counter {
    public void Bump() { value++; }
    public int Read() { return value; }
}
var counter = new Counter();
counter.Bump();
counter.Bump();
__P((counter.Read()).ToString());
__Check("2");

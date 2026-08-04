// vybe-test: csharp/csharp_default_interface_methods/default_interface_method_visible_through_interface_typed_reference
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods.rs

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

interface ICounter {
    int Value { get; }
    int Next() { return Value + 1; }
}
class Counter : ICounter {
    public int Value { get; set; }
}
ICounter counter = new Counter { Value = 4 };
__P((counter.Next()).ToString());
__Check("5");

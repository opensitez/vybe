// vybe-test: csharp/csharp_generics_constraints/generic_method_with_multiple_constraints_uses_interface_and_constructor
// origin: languages/csharp/tests/csharp/test_csharp_generics_constraints.rs

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

interface IValue { int Read(); } class Item : IValue { public int Read() { return 4; } } int Build<T>() where T : IValue, new() { return new T().Read(); } __P((Build<Item>()).ToString());
__Check("4");

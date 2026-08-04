// vybe-test: csharp/csharp_generics_constraints/generic_interface_implementation_preserves_type_argument
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

interface IBox<T> { T Read(); } class NumberBox : IBox<int> { public int Read() { return 8; } } __P((((IBox<int>)new NumberBox()).Read()).ToString());
__Check("8");

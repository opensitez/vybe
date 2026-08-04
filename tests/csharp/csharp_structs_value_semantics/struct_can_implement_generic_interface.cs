// vybe-test: csharp/csharp_structs_value_semantics/struct_can_implement_generic_interface
// origin: languages/csharp/tests/csharp/test_csharp_structs_value_semantics.rs

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

interface IBox<T> { T Read(); } struct NumberBox : IBox<int> { public int Read() { return 14; } } IBox<int> box = new NumberBox(); __P((box.Read()).ToString());
__Check("14");

// vybe-test: csharp/csharp_struct_features/readonly_struct_field_cannot_be_mutated_but_is_readable
// origin: languages/csharp/tests/csharp/test_csharp_struct_features.rs

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

readonly struct Immutable { public readonly int Value; public Immutable(int v) { Value=v; } }
var obj = new Immutable(7);
__P((obj.Value).ToString());
__Check("7");

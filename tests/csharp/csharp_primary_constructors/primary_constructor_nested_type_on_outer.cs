// vybe-test: csharp/csharp_primary_constructors/primary_constructor_nested_type_on_outer
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

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

class Outer(int seed) {
    public class Inner { public int Value; }
    public Inner Make() => new Inner { Value = seed };
}
__P((new Outer(6).Make().Value).ToString());
__Check("6");

// vybe-test: csharp/csharp_nested_partial_types/nested_class_reads_outer_constant_value
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

class Outer {
    public const string Prefix = "outer";
    public class Inner {
        public string Read() { return Prefix + "/inner"; }
    }
}
__P((new Outer.Inner().Read()).ToString());
__Check("outer/inner");

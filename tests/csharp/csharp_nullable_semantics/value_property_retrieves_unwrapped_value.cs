// vybe-test: csharp/csharp_nullable_semantics/value_property_retrieves_unwrapped_value
// origin: languages/csharp/tests/csharp/test_csharp_nullable_semantics.rs

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

int? n = 42; __P((n.Value).ToString());
__Check("42");

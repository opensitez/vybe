// vybe-test: csharp/csharp_nullable_value_type_arithmetic/nullable_conversion_from_literal_wraps_value_type
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_type_arithmetic.rs

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

int? boxed = 42;
__P((boxed is int).ToString());
__P(((int)boxed).ToString());
__Check("True\n42");

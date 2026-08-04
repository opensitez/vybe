// vybe-test: csharp/csharp_nullable_value_operators/nullable_value_operators_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_operators.rs

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

// nullable_value_operators
double seed = 57; __P(((seed + 0.5 - 0.5) == seed).ToString());
__Check("True");

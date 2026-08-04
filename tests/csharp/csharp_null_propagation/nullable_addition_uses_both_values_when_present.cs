// vybe-test: csharp/csharp_null_propagation/nullable_addition_uses_both_values_when_present
// origin: languages/csharp/tests/csharp/test_csharp_null_propagation.rs

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

int? left = 2; int? right = 5; __P(((left ?? 0) + (right ?? 0)).ToString());
__Check("7");

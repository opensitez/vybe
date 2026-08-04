// vybe-test: csharp/csharp_nullable_value_type_arithmetic/nullable_equality_compares_values_not_references
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

int? a = 7;
int? b = 7;
__P((a == b).ToString());
int? c = null;
__P((a == c).ToString());
__P((c == null).ToString());
__Check("True\nFalse\nTrue");

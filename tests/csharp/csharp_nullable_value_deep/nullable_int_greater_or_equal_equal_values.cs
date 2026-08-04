// vybe-test: csharp/csharp_nullable_value_deep/nullable_int_greater_or_equal_equal_values
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_deep.rs

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

int? a=4; int? b=4; __P((a>=b).ToString());
__Check("True");

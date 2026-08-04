// vybe-test: csharp/csharp_nullable_value_deep/nullable_int_get_value_or_default_returns_value_when_present
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

int? n=15; __P((n.GetValueOrDefault(88)).ToString());
__Check("15");

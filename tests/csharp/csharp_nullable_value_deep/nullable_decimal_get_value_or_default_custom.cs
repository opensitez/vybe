// vybe-test: csharp/csharp_nullable_value_deep/nullable_decimal_get_value_or_default_custom
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

decimal? d=null; __P((d.GetValueOrDefault(9.99m)).ToString());
__Check("9.99");

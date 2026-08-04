// vybe-test: csharp/csharp_string_parsing/int_try_parse_returns_false_for_non_numeric
// origin: languages/csharp/tests/csharp/test_csharp_string_parsing.rs

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

__P((int.TryParse("abc",out _)).ToString());
__Check("False");

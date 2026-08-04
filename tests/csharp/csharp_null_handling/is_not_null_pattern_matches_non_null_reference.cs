// vybe-test: csharp/csharp_null_handling/is_not_null_pattern_matches_non_null_reference
// origin: languages/csharp/tests/csharp/test_csharp_null_handling.rs

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

object o = "hello";
__P((o is not null).ToString());
__Check("True");

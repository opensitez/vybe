// vybe-test: csharp/csharp_string_comparison/ordinal_ignore_case_treats_same_letters_as_equal
// origin: languages/csharp/tests/csharp/test_csharp_string_comparison.rs

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

__P((string.Compare("Hello","hello",System.StringComparison.OrdinalIgnoreCase) == 0).ToString());
__Check("True");

// vybe-test: csharp/csharp_string_comparison/ordinal_comparison_is_case_sensitive_by_default
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

__P((string.Compare("A","a",System.StringComparison.Ordinal) != 0).ToString());
__Check("True");

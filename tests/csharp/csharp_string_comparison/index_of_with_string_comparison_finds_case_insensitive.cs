// vybe-test: csharp/csharp_string_comparison/index_of_with_string_comparison_finds_case_insensitive
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

__P(("fooBAR".IndexOf("bar",System.StringComparison.OrdinalIgnoreCase)).ToString());
__Check("3");

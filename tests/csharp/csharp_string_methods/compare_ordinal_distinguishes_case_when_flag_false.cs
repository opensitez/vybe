// vybe-test: csharp/csharp_string_methods/compare_ordinal_distinguishes_case_when_flag_false
// origin: languages/csharp/tests/csharp/test_csharp_string_methods.rs

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

int r = string.Compare("A","a",System.StringComparison.Ordinal);
__P((r < 0 || r > 0).ToString());
__Check("True");

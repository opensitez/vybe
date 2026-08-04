// vybe-test: csharp/csharp_verbatim_string_literals/regular_string_escape_newline_differs_from_verbatim_multiline
// origin: languages/csharp/tests/csharp/test_csharp_verbatim_string_literals.rs

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

__P(("a\nb").ToString());
__Check("a\nb");

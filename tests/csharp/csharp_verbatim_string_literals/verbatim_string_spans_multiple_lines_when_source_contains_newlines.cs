// vybe-test: csharp/csharp_verbatim_string_literals/verbatim_string_spans_multiple_lines_when_source_contains_newlines
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

__P((@"line1
line2").ToString());
__Check("line1\nline2");

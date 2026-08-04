// vybe-test: csharp/csharp_interpolated_strings/interpolated_string_multiple_holes_preserve_middle_literal_text
// origin: languages/csharp/tests/csharp/test_csharp_interpolated_strings.rs

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

__P(($"a{1}b{2}c").ToString());
__Check("a1b2c");

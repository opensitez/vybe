// vybe-test: csharp/csharp_string_interpolation/alignment_specifier_pads_right_aligned
// origin: languages/csharp/tests/csharp/test_csharp_string_interpolation.rs

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

__P(($"{"x",5}").ToString());
__Check("    x");

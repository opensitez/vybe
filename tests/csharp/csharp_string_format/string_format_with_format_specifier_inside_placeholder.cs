// vybe-test: csharp/csharp_string_format/string_format_with_format_specifier_inside_placeholder
// origin: languages/csharp/tests/csharp/test_csharp_string_format.rs

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

__P((string.Format("{0:F1}", 3.14159)).ToString());
__Check("3.1");

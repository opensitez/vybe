// vybe-test: csharp/csharp_numeric_formatting/format_x_lower_encodes_integer_as_lowercase_hex
// origin: languages/csharp/tests/csharp/test_csharp_numeric_formatting.rs

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

__P((255.ToString("x")).ToString());
__Check("ff");

// vybe-test: csharp/csharp_utf8_string_literals/utf8_literal_empty_has_zero_length
// origin: languages/csharp/tests/csharp/test_csharp_utf8_string_literals.rs

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

var bytes=u8""; __P((bytes.Length).ToString());
__Check("0");

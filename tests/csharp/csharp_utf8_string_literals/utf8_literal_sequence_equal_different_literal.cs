// vybe-test: csharp/csharp_utf8_string_literals/utf8_literal_sequence_equal_different_literal
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

var a=u8"one"; var b=u8"two"; __P((a.SequenceEqual(b)).ToString());
__Check("False");

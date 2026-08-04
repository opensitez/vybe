// vybe-test: csharp/csharp_encoding/unicode_encoding_uses_two_bytes_per_char
// origin: languages/csharp/tests/csharp/test_csharp_encoding.rs

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

var bytes = System.Text.Encoding.Unicode.GetBytes("A");
__P((bytes.Length).ToString());
__Check("2");

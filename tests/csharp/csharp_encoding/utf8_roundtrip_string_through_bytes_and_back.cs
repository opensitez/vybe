// vybe-test: csharp/csharp_encoding/utf8_roundtrip_string_through_bytes_and_back
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

var bytes = System.Text.Encoding.UTF8.GetBytes("hello");
__P((System.Text.Encoding.UTF8.GetString(bytes)).ToString());
__Check("hello");

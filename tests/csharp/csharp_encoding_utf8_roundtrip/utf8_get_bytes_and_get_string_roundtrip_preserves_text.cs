// vybe-test: csharp/csharp_encoding_utf8_roundtrip/utf8_get_bytes_and_get_string_roundtrip_preserves_text
// origin: languages/csharp/tests/csharp/test_csharp_encoding_utf8_roundtrip.rs

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

var encoding = System.Text.Encoding.UTF8;
var bytes = encoding.GetBytes("café");
var text = encoding.GetString(bytes);
__P((text).ToString());
__Check("café");

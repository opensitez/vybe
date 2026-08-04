// vybe-test: csharp/csharp_encoding/convert_between_encodings_via_encode_decode_roundtrip
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

string text = "test";
byte[] bytes = System.Text.Encoding.UTF8.GetBytes(text);
string result = System.Text.Encoding.UTF8.GetString(bytes);
__P((text == result).ToString());
__Check("True");

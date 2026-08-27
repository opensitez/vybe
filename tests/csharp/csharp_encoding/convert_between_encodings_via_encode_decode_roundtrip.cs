// vybe-test: csharp/csharp_encoding/convert_between_encodings_via_encode_decode_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_encoding.rs

using static __Harness;

string text = "test";
byte[] bytes = System.Text.Encoding.UTF8.GetBytes(text);
string result = System.Text.Encoding.UTF8.GetString(bytes);
__P((text == result).ToString());
__Check("True");

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}

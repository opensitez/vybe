// vybe-test: csharp/csharp_convert_uri_path/encoding_utf8_roundtrips_unicode_text
// origin: languages/csharp/tests/csharp/test_csharp_convert_uri_path.rs

using static __Harness;

var bytes = System.Text.Encoding.UTF8.GetBytes("café");
var text = System.Text.Encoding.UTF8.GetString(bytes);
__P((text).ToString());
__Check("café");

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

// vybe-test: csharp/csharp_convert_uri_path/convert_from_base64_roundtrips_byte_payload
// origin: languages/csharp/tests/csharp/test_csharp_convert_uri_path.rs

using static __Harness;

var bytes = System.Convert.FromBase64String("AQID");
__P((bytes.Length).ToString());
__P((bytes[2]).ToString());
__Check("3\n3");

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

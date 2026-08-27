// vybe-test: csharp/csharp_encoding/get_byte_count_reflects_character_byte_width
// origin: languages/csharp/tests/csharp/test_csharp_encoding.rs

using static __Harness;

int n = System.Text.Encoding.UTF8.GetByteCount("café");
__P((n > 4).ToString());
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

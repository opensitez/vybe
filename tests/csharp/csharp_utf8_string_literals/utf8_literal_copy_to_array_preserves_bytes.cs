// vybe-test: csharp/csharp_utf8_string_literals/utf8_literal_copy_to_array_preserves_bytes
// origin: languages/csharp/tests/csharp/test_csharp_utf8_string_literals.rs

using static __Harness;

var bytes="xy"u8;
byte[] buf=new byte[2];
bytes.CopyTo(buf);
__P((buf[0]).ToString());
__P((buf[1]).ToString());
__Check("120\n121");

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

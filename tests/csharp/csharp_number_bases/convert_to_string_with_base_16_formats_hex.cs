// vybe-test: csharp/csharp_number_bases/convert_to_string_with_base_16_formats_hex
// origin: languages/csharp/tests/csharp/test_csharp_number_bases.rs

using static __Harness;

__P((System.Convert.ToString(255,16)).ToString());
__Check("ff");

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

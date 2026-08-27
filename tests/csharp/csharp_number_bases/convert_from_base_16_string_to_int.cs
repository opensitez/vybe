// vybe-test: csharp/csharp_number_bases/convert_from_base_16_string_to_int
// origin: languages/csharp/tests/csharp/test_csharp_number_bases.rs

using static __Harness;

__P((System.Convert.ToInt32("ff",16)).ToString());
__Check("255");

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

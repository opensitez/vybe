// vybe-test: csharp/csharp_type_conversions/char_to_integer_cast_produces_code_point
// origin: languages/csharp/tests/csharp/test_csharp_type_conversions.rs

using static __Harness;

char ch = 'A';
__P(((int)ch).ToString());
__Check("65");

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

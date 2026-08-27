// vybe-test: csharp/strings_advanced/int_parse_tostring
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

using static __Harness;

int x = int.Parse("42");
__P((x + 8).ToString());
__P((x.ToString()).ToString());
__Check("50\n42");

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

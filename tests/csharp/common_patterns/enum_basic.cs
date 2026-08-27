// vybe-test: csharp/common_patterns/enum_basic
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

using static __Harness;

Color c = Color.Green;
__P((c).ToString());
__P(((int)c).ToString());
__Check("Green\n1");

enum Color { Red, Green, Blue }

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

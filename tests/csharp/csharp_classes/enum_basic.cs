// vybe-test: csharp/csharp_classes/enum_basic
// origin: languages/csharp/tests/csharp/test_csharp_classes.rs

using static __Harness;

Color c = Color.Green;
__P(((int)c).ToString());
__Check("1");

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

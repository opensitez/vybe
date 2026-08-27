// vybe-test: csharp/strings_advanced/stringbuilder_insert_replace
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

using static __Harness;

var sb = new System.Text.StringBuilder("Hello World");
sb.Replace("World", "There");
__P((sb.ToString()).ToString());
sb.Insert(5, " Beautiful");
__P((sb.ToString()).ToString());
__Check("Hello There\nHello Beautiful There");

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

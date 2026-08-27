// vybe-test: csharp/strings_advanced/string_insert_remove
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

using static __Harness;

string s = "Hello World";
__P((s.Insert(5, " Beautiful")).ToString());
__P((s.Remove(5)).ToString());
__P((s.Remove(5, 1)).ToString());
__Check("Hello Beautiful World\nHello\nHelloWorld");

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

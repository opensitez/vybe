// vybe-test: csharp/strings_advanced/string_format
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

using static __Harness;

__P((string.Format("{0} + {1} = {2}", 1, 2, 3)).ToString());
__P((string.Format("Name: {0}, Age: {1}", "Bob", 25)).ToString());
__Check("1 + 2 = 3\nName: Bob, Age: 25");

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

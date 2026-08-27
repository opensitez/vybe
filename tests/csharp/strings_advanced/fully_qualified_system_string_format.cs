// vybe-test: csharp/strings_advanced/fully_qualified_system_string_format
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

using static __Harness;

__P((System.String.Format("{0}-{1}", "A", "B")).ToString());
__Check("A-B");

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

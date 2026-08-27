// vybe-test: csharp/csharp_exception_types/format_exception_thrown_by_parse_on_non_numeric_string
// origin: languages/csharp/tests/csharp/test_csharp_exception_types.rs

using static __Harness;

string result = "";
try { int.Parse("abc"); }
catch(System.FormatException) { result = "fmt"; }
__P((result).ToString());
__Check("fmt");

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

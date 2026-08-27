// vybe-test: csharp/csharp_strings_ext/verbatim_string
// origin: languages/csharp/tests/csharp/test_csharp_strings_ext.rs

using static __Harness;

var path = @"C:\Users\test\file.txt";
__P((path).ToString());
__Check("C:\\Users\\test\\file.txt");

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

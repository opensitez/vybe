// vybe-test: csharp/csharp_string_raw_verbatim/verbatim_string_preserves_backslashes
// origin: languages/csharp/tests/csharp/test_csharp_string_raw_verbatim.rs

using static __Harness;

string path=@"C:\Users\test\file.txt";
__P((path.Contains(@"\test")).ToString());
__Check("True");

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

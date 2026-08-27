// vybe-test: csharp/csharp_string_interpolation/verbatim_interpolated_string_preserves_backslash
// origin: languages/csharp/tests/csharp/test_csharp_string_interpolation.rs

using static __Harness;

string dir="docs";
__P(($@"C:\{dir}\file.txt").ToString());
__Check("C:\\docs\\file.txt");

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

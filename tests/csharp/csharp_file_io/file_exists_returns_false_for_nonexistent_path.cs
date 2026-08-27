// vybe-test: csharp/csharp_file_io/file_exists_returns_false_for_nonexistent_path
// origin: languages/csharp/tests/csharp/test_csharp_file_io.rs

using static __Harness;

__P((System.IO.File.Exists("/no/such/path/xyz123.txt")).ToString());
__Check("False");

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

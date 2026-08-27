// vybe-test: csharp/csharp_directory_io/get_temp_path_returns_non_empty_string
// origin: languages/csharp/tests/csharp/test_csharp_directory_io.rs

using static __Harness;

__P((System.IO.Path.GetTempPath().Length > 0).ToString());
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

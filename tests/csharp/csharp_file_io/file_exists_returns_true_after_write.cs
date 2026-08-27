// vybe-test: csharp/csharp_file_io/file_exists_returns_true_after_write
// origin: languages/csharp/tests/csharp/test_csharp_file_io.rs

using static __Harness;

string path = System.IO.Path.GetTempFileName();
__P((System.IO.File.Exists(path)).ToString());
System.IO.File.Delete(path);
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

// vybe-test: csharp/csharp_directory_io/path_get_directory_name_returns_parent_path
// origin: languages/csharp/tests/csharp/test_csharp_directory_io.rs

using static __Harness;

string dir = System.IO.Path.GetDirectoryName("/tmp/file.txt");
__P((dir).ToString());
__Check("/tmp");

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

// vybe-test: csharp/csharp_convert_uri_path/path_get_extension_returns_suffix_including_dot
// origin: languages/csharp/tests/csharp/test_csharp_convert_uri_path.rs

using static __Harness;

// convert_uri_path
__P((System.IO.Path.GetExtension("archive.tar.gz")).ToString());
__Check(".gz");

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

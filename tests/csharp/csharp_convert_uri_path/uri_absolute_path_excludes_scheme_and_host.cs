// vybe-test: csharp/csharp_convert_uri_path/uri_absolute_path_excludes_scheme_and_host
// origin: languages/csharp/tests/csharp/test_csharp_convert_uri_path.rs

using static __Harness;

var link = new System.Uri("https://example.com/api/v1");
__P((link.AbsolutePath).ToString());
__Check("/api/v1");

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

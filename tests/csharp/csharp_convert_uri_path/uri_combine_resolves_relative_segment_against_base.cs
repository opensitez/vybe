// vybe-test: csharp/csharp_convert_uri_path/uri_combine_resolves_relative_segment_against_base
// origin: languages/csharp/tests/csharp/test_csharp_convert_uri_path.rs

using static __Harness;

var baseUri = new System.Uri("https://example.com/a/");
var combined = new System.Uri(baseUri, "b");
__P((combined.AbsolutePath).ToString());
__Check("/a/b");

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

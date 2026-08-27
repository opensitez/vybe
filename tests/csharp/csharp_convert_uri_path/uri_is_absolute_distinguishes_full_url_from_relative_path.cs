// vybe-test: csharp/csharp_convert_uri_path/uri_is_absolute_distinguishes_full_url_from_relative_path
// origin: languages/csharp/tests/csharp/test_csharp_convert_uri_path.rs

using static __Harness;

var absolute = new System.Uri("https://example.com");
var relative = new System.Uri("/only-path", System.UriKind.Relative);
__P((absolute.IsAbsoluteUri).ToString());
__P((relative.IsAbsoluteUri).ToString());
__Check("True\nFalse");

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

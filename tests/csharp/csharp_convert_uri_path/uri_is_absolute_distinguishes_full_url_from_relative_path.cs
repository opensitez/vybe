// vybe-test: csharp/csharp_convert_uri_path/uri_is_absolute_distinguishes_full_url_from_relative_path
// origin: languages/csharp/tests/csharp/test_csharp_convert_uri_path.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

var absolute = new System.Uri("https://example.com");
var relative = new System.Uri("/only-path", System.UriKind.Relative);
__P((absolute.IsAbsoluteUri).ToString());
__P((relative.IsAbsoluteUri).ToString());
__Check("True\nFalse");

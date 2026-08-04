// vybe-test: csharp/csharp_convert_uri_path/uri_absolute_path_excludes_scheme_and_host
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

var link = new System.Uri("https://example.com/api/v1");
__P((link.AbsolutePath).ToString());
__Check("/api/v1");

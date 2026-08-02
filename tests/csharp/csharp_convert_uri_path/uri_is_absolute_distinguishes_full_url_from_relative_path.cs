// vybe-test: csharp/csharp_convert_uri_path/uri_is_absolute_distinguishes_full_url_from_relative_path
// origin: languages/csharp/tests/csharp/test_csharp_convert_uri_path.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var absolute = new System.Uri("https://example.com");
var relative = new System.Uri("/only-path", System.UriKind.Relative);
__Check((absolute.IsAbsoluteUri).ToString(), "True");
__Check((relative.IsAbsoluteUri).ToString(), "False");

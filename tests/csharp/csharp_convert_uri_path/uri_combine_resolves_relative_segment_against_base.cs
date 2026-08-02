// vybe-test: csharp/csharp_convert_uri_path/uri_combine_resolves_relative_segment_against_base
// origin: languages/csharp/tests/csharp/test_csharp_convert_uri_path.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var baseUri = new System.Uri("https://example.com/a/");
var combined = new System.Uri(baseUri, "b");
__Check((combined.AbsolutePath).ToString(), "/a/b");

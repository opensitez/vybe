// vybe-test: csharp/csharp_convert_uri_path/uri_absolute_path_excludes_scheme_and_host
// origin: languages/csharp/tests/csharp/test_csharp_convert_uri_path.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var link = new System.Uri("https://example.com/api/v1");
__Check((link.AbsolutePath).ToString(), "/api/v1");

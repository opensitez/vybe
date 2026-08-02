// vybe-test: csharp/csharp_convert_uri_path/path_get_extension_returns_suffix_including_dot
// origin: languages/csharp/tests/csharp/test_csharp_convert_uri_path.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// convert_uri_path
__Check((System.IO.Path.GetExtension("archive.tar.gz")).ToString(), ".gz");

// vybe-test: csharp/csharp_convert_uri_path/path_change_extension_replaces_trailing_suffix
// origin: languages/csharp/tests/csharp/test_csharp_convert_uri_path.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// convert_uri_path
__Check((System.IO.Path.ChangeExtension("data.txt", ".json")).ToString(), "data.json");

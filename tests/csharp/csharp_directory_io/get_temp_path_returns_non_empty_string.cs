// vybe-test: csharp/csharp_directory_io/get_temp_path_returns_non_empty_string
// origin: languages/csharp/tests/csharp/test_csharp_directory_io.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.IO.Path.GetTempPath().Length > 0).ToString(), "True");

// vybe-test: csharp/csharp_file_io/file_exists_returns_false_for_nonexistent_path
// origin: languages/csharp/tests/csharp/test_csharp_file_io.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.IO.File.Exists("/no/such/path/xyz123.txt")).ToString(), "False");

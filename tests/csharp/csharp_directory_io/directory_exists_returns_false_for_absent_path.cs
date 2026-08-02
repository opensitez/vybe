// vybe-test: csharp/csharp_directory_io/directory_exists_returns_false_for_absent_path
// origin: languages/csharp/tests/csharp/test_csharp_directory_io.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.IO.Directory.Exists("/no/such/dir/xyz999")).ToString(), "False");

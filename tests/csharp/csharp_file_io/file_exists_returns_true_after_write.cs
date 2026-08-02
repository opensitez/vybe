// vybe-test: csharp/csharp_file_io/file_exists_returns_true_after_write
// origin: languages/csharp/tests/csharp/test_csharp_file_io.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string path = System.IO.Path.GetTempFileName();
__Check((System.IO.File.Exists(path)).ToString(), "True");
System.IO.File.Delete(path);

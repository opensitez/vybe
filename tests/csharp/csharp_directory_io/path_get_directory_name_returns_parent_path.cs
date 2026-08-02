// vybe-test: csharp/csharp_directory_io/path_get_directory_name_returns_parent_path
// origin: languages/csharp/tests/csharp/test_csharp_directory_io.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string dir = System.IO.Path.GetDirectoryName("/tmp/file.txt");
__Check((dir).ToString(), "/tmp");

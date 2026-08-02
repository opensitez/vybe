// vybe-test: csharp/csharp_directory_io/directory_create_makes_new_folder
// origin: languages/csharp/tests/csharp/test_csharp_directory_io.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "vybe_test_"+System.Guid.NewGuid().ToString("N"));
System.IO.Directory.CreateDirectory(path);
__Check((System.IO.Directory.Exists(path)).ToString(), "True");
System.IO.Directory.Delete(path);

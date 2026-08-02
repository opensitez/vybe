// vybe-test: csharp/csharp_file_io/append_all_text_adds_to_existing_file
// origin: languages/csharp/tests/csharp/test_csharp_file_io.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string path = System.IO.Path.GetTempFileName();
System.IO.File.WriteAllText(path, "hello");
System.IO.File.AppendAllText(path, " world");
__Check((System.IO.File.ReadAllText(path)).ToString(), "hello world");
System.IO.File.Delete(path);

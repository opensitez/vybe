// vybe-test: csharp/csharp_file_io/write_all_text_then_read_all_text_roundtrips
// origin: languages/csharp/tests/csharp/test_csharp_file_io.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string path = System.IO.Path.GetTempFileName();
System.IO.File.WriteAllText(path, "hello");
__Check((System.IO.File.ReadAllText(path)).ToString(), "hello");
System.IO.File.Delete(path);

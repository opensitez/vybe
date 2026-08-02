// vybe-test: csharp/csharp_file_io/read_all_bytes_count_matches_written_byte_length
// origin: languages/csharp/tests/csharp/test_csharp_file_io.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string path = System.IO.Path.GetTempFileName();
System.IO.File.WriteAllBytes(path, new byte[]{1,2,3,4,5});
__Check((System.IO.File.ReadAllBytes(path).Length).ToString(), "5");
System.IO.File.Delete(path);

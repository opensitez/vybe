// vybe-test: csharp/csharp_file_io/write_all_lines_then_read_all_lines_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_file_io.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string path = System.IO.Path.GetTempFileName();
System.IO.File.WriteAllLines(path, new[]{"a","b","c"});
var lines = System.IO.File.ReadAllLines(path);
__Check((lines.Length).ToString(), "3");
__Check((lines[1]).ToString(), "b");
System.IO.File.Delete(path);

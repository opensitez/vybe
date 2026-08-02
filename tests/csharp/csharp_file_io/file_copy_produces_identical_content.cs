// vybe-test: csharp/csharp_file_io/file_copy_produces_identical_content
// origin: languages/csharp/tests/csharp/test_csharp_file_io.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string src = System.IO.Path.GetTempFileName();
string dst = src + ".copy";
System.IO.File.WriteAllText(src, "data");
System.IO.File.Copy(src, dst, true);
__Check((System.IO.File.ReadAllText(dst)).ToString(), "data");
System.IO.File.Delete(src);
System.IO.File.Delete(dst);

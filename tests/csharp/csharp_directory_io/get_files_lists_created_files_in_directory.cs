// vybe-test: csharp/csharp_directory_io/get_files_lists_created_files_in_directory
// origin: languages/csharp/tests/csharp/test_csharp_directory_io.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string dir = System.IO.Path.Combine(System.IO.Path.GetTempPath(),"vybe_"+System.Guid.NewGuid().ToString("N"));
System.IO.Directory.CreateDirectory(dir);
System.IO.File.WriteAllText(System.IO.Path.Combine(dir,"a.txt"),"a");
System.IO.File.WriteAllText(System.IO.Path.Combine(dir,"b.txt"),"b");
__Check((System.IO.Directory.GetFiles(dir).Length).ToString(), "2");
System.IO.Directory.Delete(dir, true);

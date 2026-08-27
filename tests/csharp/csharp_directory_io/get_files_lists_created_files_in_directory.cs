// vybe-test: csharp/csharp_directory_io/get_files_lists_created_files_in_directory
// origin: languages/csharp/tests/csharp/test_csharp_directory_io.rs

using static __Harness;

string dir = System.IO.Path.Combine(System.IO.Path.GetTempPath(),"vybe_"+System.Guid.NewGuid().ToString("N"));
System.IO.Directory.CreateDirectory(dir);
System.IO.File.WriteAllText(System.IO.Path.Combine(dir,"a.txt"),"a");
System.IO.File.WriteAllText(System.IO.Path.Combine(dir,"b.txt"),"b");
__P((System.IO.Directory.GetFiles(dir).Length).ToString());
System.IO.Directory.Delete(dir, true);
__Check("2");

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}

// vybe-test: csharp/csharp_directory_io/directory_create_makes_new_folder
// origin: languages/csharp/tests/csharp/test_csharp_directory_io.rs

using static __Harness;

string path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "vybe_test_"+System.Guid.NewGuid().ToString("N"));
System.IO.Directory.CreateDirectory(path);
__P((System.IO.Directory.Exists(path)).ToString());
System.IO.Directory.Delete(path);
__Check("True");

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

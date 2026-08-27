// vybe-test: csharp/csharp_file_io/file_copy_produces_identical_content
// origin: languages/csharp/tests/csharp/test_csharp_file_io.rs

using static __Harness;

string src = System.IO.Path.GetTempFileName();
string dst = src + ".copy";
System.IO.File.WriteAllText(src, "data");
System.IO.File.Copy(src, dst, true);
__P((System.IO.File.ReadAllText(dst)).ToString());
System.IO.File.Delete(src);
System.IO.File.Delete(dst);
__Check("data");

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

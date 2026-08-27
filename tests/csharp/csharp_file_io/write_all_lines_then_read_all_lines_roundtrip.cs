// vybe-test: csharp/csharp_file_io/write_all_lines_then_read_all_lines_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_file_io.rs

using static __Harness;

string path = System.IO.Path.GetTempFileName();
System.IO.File.WriteAllLines(path, new[]{"a","b","c"});
var lines = System.IO.File.ReadAllLines(path);
__P((lines.Length).ToString());
__P((lines[1]).ToString());
System.IO.File.Delete(path);
__Check("3\nb");

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

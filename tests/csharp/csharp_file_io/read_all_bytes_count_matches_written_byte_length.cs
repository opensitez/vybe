// vybe-test: csharp/csharp_file_io/read_all_bytes_count_matches_written_byte_length
// origin: languages/csharp/tests/csharp/test_csharp_file_io.rs

using static __Harness;

string path = System.IO.Path.GetTempFileName();
System.IO.File.WriteAllBytes(path, new byte[]{1,2,3,4,5});
__P((System.IO.File.ReadAllBytes(path).Length).ToString());
System.IO.File.Delete(path);
__Check("5");

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

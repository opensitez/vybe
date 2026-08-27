// vybe-test: csharp/csharp_io_file_read_surface/io_file_read_surface_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_io_file_read_surface.rs

using static __Harness;

// io_file_read_surface
string feature = "io_file_read_surface:89";
__P((feature.Length >= 1).ToString());
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

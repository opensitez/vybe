// vybe-test: csharp/csharp_records_advanced/readonly_record_struct_exposes_members
// origin: languages/csharp/tests/csharp/test_csharp_records_advanced.rs

using static __Harness;

var size = new Size(4, 6);
__P((size.Width * size.Height).ToString());
__Check("24");

readonly record struct Size(int Width, int Height);

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

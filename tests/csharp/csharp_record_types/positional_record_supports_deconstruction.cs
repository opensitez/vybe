// vybe-test: csharp/csharp_record_types/positional_record_supports_deconstruction
// origin: languages/csharp/tests/csharp/test_csharp_record_types.rs

using static __Harness;

var s = new Size(10,20);
var (w,h) = s;
__P((w).ToString());
__P((h).ToString());
__Check("10\n20");

record Size(int W, int H);

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

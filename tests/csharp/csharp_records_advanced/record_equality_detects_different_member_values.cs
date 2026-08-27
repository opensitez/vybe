// vybe-test: csharp/csharp_records_advanced/record_equality_detects_different_member_values
// origin: languages/csharp/tests/csharp/test_csharp_records_advanced.rs

using static __Harness;

__P((new Point(1, 2) == new Point(2, 1)).ToString());
__Check("False");

record Point(int X, int Y);

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

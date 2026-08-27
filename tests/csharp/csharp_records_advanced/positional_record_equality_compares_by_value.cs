// vybe-test: csharp/csharp_records_advanced/positional_record_equality_compares_by_value
// origin: languages/csharp/tests/csharp/test_csharp_records_advanced.rs

using static __Harness;

__P((new Point(1, 2) == new Point(1, 2)).ToString());
__Check("True");

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

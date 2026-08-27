// vybe-test: csharp/csharp_pattern_deconstruct/positional_pattern_matches_deconstructed_record_fields
// origin: languages/csharp/tests/csharp/test_csharp_pattern_deconstruct.rs

using static __Harness;

object obj = new Point(0, 5);
if (obj is Point(0, var y)) __P((y).ToString());
else __P((-1).ToString());
__Check("5");

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

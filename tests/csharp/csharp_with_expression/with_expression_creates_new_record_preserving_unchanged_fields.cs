// vybe-test: csharp/csharp_with_expression/with_expression_creates_new_record_preserving_unchanged_fields
// origin: languages/csharp/tests/csharp/test_csharp_with_expression.rs

using static __Harness;

var origin = new Point(1, 2);
var moved = origin with { X = 10 }
;
__P((moved.X).ToString());
__P((moved.Y).ToString());
__Check("10\n2");

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

// vybe-test: csharp/csharp_with_expression_records/with_positional_single_field
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

using static __Harness;

var q=(new Point(1,2)) with{X=10}
;
__P((q.X).ToString());
__P((q.Y).ToString());
__Check("10\n2");

record Point(int X,int Y);

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

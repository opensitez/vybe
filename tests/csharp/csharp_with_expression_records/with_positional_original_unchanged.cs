// vybe-test: csharp/csharp_with_expression_records/with_positional_original_unchanged
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

using static __Harness;

var p=new Point(1,2);
var q=p with{X=9}
;
__P((p.X).ToString());
__Check("1");

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

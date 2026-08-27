// vybe-test: csharp/csharp_with_expression_records/with_four_positional
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

using static __Harness;

var r=(new Quad(1,2,3,4)) with{D=10}
;
__P((r.D).ToString());
__Check("10");

record Quad(int A,int B,int C,int D);

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

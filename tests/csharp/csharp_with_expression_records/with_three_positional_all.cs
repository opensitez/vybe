// vybe-test: csharp/csharp_with_expression_records/with_three_positional_all
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

using static __Harness;

var u=(new Triple(1,2,3)) with{A=4,B=5,C=6}
;
__P((u.A+u.B+u.C).ToString());
__Check("15");

record Triple(int A,int B,int C);

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

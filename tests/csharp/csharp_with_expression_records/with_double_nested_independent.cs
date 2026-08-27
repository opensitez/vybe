// vybe-test: csharp/csharp_with_expression_records/with_double_nested_independent
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

using static __Harness;

var p=new Pair(1,1);
var a=p with{A=2}
;
var b=p with{B=3}
;
__P((a.A).ToString());
__P((b.B).ToString());
__Check("2\n3");

record Pair(int A,int B);

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

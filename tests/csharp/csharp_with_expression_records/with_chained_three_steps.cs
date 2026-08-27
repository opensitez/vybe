// vybe-test: csharp/csharp_with_expression_records/with_chained_three_steps
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

using static __Harness;

var a=new Box(1,2,3);
var b=a with{W=4}
;
var c=b with{H=5}
;
var d=c with{D=6}
;
__P((a.W).ToString());
__P((d.W).ToString());
__P((d.H).ToString());
__P((d.D).ToString());
__Check("1\n4\n5\n6");

record Box(int W,int H,int D);

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

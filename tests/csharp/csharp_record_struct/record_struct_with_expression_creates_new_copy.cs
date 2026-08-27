// vybe-test: csharp/csharp_record_struct/record_struct_with_expression_creates_new_copy
// origin: languages/csharp/tests/csharp/test_csharp_record_struct.rs

using static __Harness;

var a=new Point(1,2);
var b=a with{X=99}
;
__P((a.X).ToString());
__P((b.X).ToString());
__Check("1\n99");

record struct Point(int X,int Y);

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

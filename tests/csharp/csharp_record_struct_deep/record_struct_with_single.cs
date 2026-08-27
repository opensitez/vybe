// vybe-test: csharp/csharp_record_struct_deep/record_struct_with_single
// origin: languages/csharp/tests/csharp/test_csharp_record_struct_deep.rs

using static __Harness;

var p=new Point(1,2);
var q=p with{X=9}
;
__P((p.X).ToString());
__P((q.X).ToString());
__Check("1\n9");

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

// vybe-test: csharp/csharp_record_struct_deep/record_struct_with_two
// origin: languages/csharp/tests/csharp/test_csharp_record_struct_deep.rs

using static __Harness;

var p=new Point(1,2);
var q=p with{X=3,Y=4}
;
__P((p.Y).ToString());
__P((q.X).ToString());
__P((q.Y).ToString());
__Check("2\n3\n4");

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

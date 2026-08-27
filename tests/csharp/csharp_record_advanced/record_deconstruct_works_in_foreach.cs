// vybe-test: csharp/csharp_record_advanced/record_deconstruct_works_in_foreach
// origin: languages/csharp/tests/csharp/test_csharp_record_advanced.rs

using static __Harness;

var pts=new[]{new Point(1,2),new Point(3,4)}
;
int sumX=0;
foreach(var(x,_) in pts) sumX+=x;
__P((sumX).ToString());
__Check("4");

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

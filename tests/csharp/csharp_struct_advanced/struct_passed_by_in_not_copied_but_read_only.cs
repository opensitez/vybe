// vybe-test: csharp/csharp_struct_advanced/struct_passed_by_in_not_copied_but_read_only
// origin: languages/csharp/tests/csharp/test_csharp_struct_advanced.rs

using static __Harness;

int Sum(in Vec v)=>v.X+v.Y;
var v=new Vec{X=3,Y=4}
;
__P((Sum(in v)).ToString());
__Check("7");

struct Vec{public int X,Y;}

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

// vybe-test: csharp/csharp_record_advanced/record_struct_has_value_semantics
// origin: languages/csharp/tests/csharp/test_csharp_record_advanced.rs

using static __Harness;

var a=new Vec(1,2);
var b=a;
// copy
b=b with{X=99}
;
__P((a.X).ToString());
__Check("1");

record struct Vec(int X,int Y);

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

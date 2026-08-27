// vybe-test: csharp/csharp_struct_advanced/struct_default_keyword_produces_zero_fields
// origin: languages/csharp/tests/csharp/test_csharp_struct_advanced.rs

using static __Harness;

var v=default(Vec);
__P((v.X==0&&v.Y==0&&v.Z==0).ToString());
__Check("True");

struct Vec{public int X,Y,Z;}

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

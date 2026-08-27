// vybe-test: csharp/csharp_operator_overloading/unary_negation_operator_flips_sign_of_fields
// origin: languages/csharp/tests/csharp/test_csharp_operator_overloading.rs

using static __Harness;

var v=-new Vec{X=7}
;
__P((v.X).ToString());
__Check("-7");

struct Vec{public int X;
public static Vec operator-(Vec v)=>new Vec{X=-v.X};}

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

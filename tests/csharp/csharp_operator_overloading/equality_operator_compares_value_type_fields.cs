// vybe-test: csharp/csharp_operator_overloading/equality_operator_compares_value_type_fields
// origin: languages/csharp/tests/csharp/test_csharp_operator_overloading.rs

using static __Harness;

var a=new Color{R=1,G=2,B=3}
;
var b=new Color{R=1,G=2,B=3}
;
__P((a==b).ToString());
__P((a!=b).ToString());
__Check("True\nFalse");

struct Color{public int R,G,B;
public static bool operator==(Color a,Color b)=>a.R==b.R&&a.G==b.G&&a.B==b.B;
public static bool operator!=(Color a,Color b)=>!(a==b);
public override int GetHashCode()=>0; public override bool Equals(object o)=>o is Color c&&c==this;}

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

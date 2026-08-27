// vybe-test: csharp/csharp_operator_overloading/comparison_operators_less_and_greater
// origin: languages/csharp/tests/csharp/test_csharp_operator_overloading.rs

using static __Harness;

var a=new Weight{Kg=5}
;
var b=new Weight{Kg=10}
;
__P((a<b).ToString());
__P((a>b).ToString());
__Check("True\nFalse");

class Weight:System.IComparable<Weight>{
public int Kg;
public static bool operator<(Weight a,Weight b)=>a.Kg<b.Kg;
public static bool operator>(Weight a,Weight b)=>a.Kg>b.Kg;
public int CompareTo(Weight o)=>Kg.CompareTo(o.Kg);}

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

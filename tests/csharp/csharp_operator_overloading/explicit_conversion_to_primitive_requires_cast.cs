// vybe-test: csharp/csharp_operator_overloading/explicit_conversion_to_primitive_requires_cast
// origin: languages/csharp/tests/csharp/test_csharp_operator_overloading.rs

using static __Harness;

var p=new Percent{Value=50}
;
__P(((double)p).ToString());
__Check("0.5");

struct Percent{public double Value;
public static explicit operator double(Percent p)=>p.Value/100.0;}

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

// vybe-test: csharp/csharp_properties_advanced/computed_property_recalculates_on_each_access
// origin: languages/csharp/tests/csharp/test_csharp_properties_advanced.rs

using static __Harness;

var c=new Circle{Radius=1.0}
;
__P((System.Math.Round(c.Circumference,5)).ToString());
__Check("6.28319");

class Circle{
    public double Radius;
    public double Circumference=>2*System.Math.PI*Radius;
}

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

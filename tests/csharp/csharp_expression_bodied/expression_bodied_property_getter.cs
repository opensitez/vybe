// vybe-test: csharp/csharp_expression_bodied/expression_bodied_property_getter
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied.rs

using static __Harness;

__P((System.Math.Round(new Circle{R=0}.Area)).ToString());
__Check("0");

class Circle{public double R;public double Area=>System.Math.PI*R*R;}

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

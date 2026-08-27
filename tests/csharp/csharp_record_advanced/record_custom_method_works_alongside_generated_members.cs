// vybe-test: csharp/csharp_record_advanced/record_custom_method_works_alongside_generated_members
// origin: languages/csharp/tests/csharp/test_csharp_record_advanced.rs

using static __Harness;

var c=new Circle(1.0);
__P((c.Area>3.1&&c.Area<3.2).ToString());
__Check("True");

record Circle(double Radius){
    public double Area=>System.Math.PI*Radius*Radius;
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

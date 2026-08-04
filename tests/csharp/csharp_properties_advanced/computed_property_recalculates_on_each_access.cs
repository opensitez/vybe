// vybe-test: csharp/csharp_properties_advanced/computed_property_recalculates_on_each_access
// origin: languages/csharp/tests/csharp/test_csharp_properties_advanced.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

class Circle{
    public double Radius;
    public double Circumference=>2*System.Math.PI*Radius;
}
var c=new Circle{Radius=1.0};
__P((System.Math.Round(c.Circumference,5)).ToString());
__Check("6.28319");

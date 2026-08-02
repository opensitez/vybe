// vybe-test: csharp/csharp_properties_advanced/computed_property_recalculates_on_each_access
// origin: languages/csharp/tests/csharp/test_csharp_properties_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Circle{
    public double Radius;
    public double Circumference=>2*System.Math.PI*Radius;
}
var c=new Circle{Radius=1.0};
__Check((System.Math.Round(c.Circumference,5)).ToString(), "6.28319");

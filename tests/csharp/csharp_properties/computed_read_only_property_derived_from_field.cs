// vybe-test: csharp/csharp_properties/computed_read_only_property_derived_from_field
// origin: languages/csharp/tests/csharp/test_csharp_properties.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Circle { public double Radius; public double Area => System.Math.PI * Radius * Radius; }
__Check((System.Math.Round(new Circle{Radius=0}.Area)).ToString(), "0");

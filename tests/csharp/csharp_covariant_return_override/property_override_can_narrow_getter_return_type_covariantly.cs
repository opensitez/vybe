// vybe-test: csharp/csharp_covariant_return_override/property_override_can_narrow_getter_return_type_covariantly
// origin: languages/csharp/tests/csharp/test_csharp_covariant_return_override.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Shape { public virtual Shape CloneShape() { return new Shape(); } }
class Circle : Shape { public int Radius; public override Circle CloneShape() { return new Circle { Radius = Radius }; } }
Shape original = new Circle { Radius = 5 };
var copy = original.CloneShape();
__Check((copy is Circle).ToString(), "True");

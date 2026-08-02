// vybe-test: csharp/csharp_virtual_dispatch_semantics/virtual_property_getter_dispatches_to_derived_accessor
// origin: languages/csharp/tests/csharp/test_csharp_virtual_dispatch_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Shape {
    public virtual int Sides { get { return 0; } }
}
class Triangle : Shape {
    public override int Sides { get { return 3; } }
}
Shape shape = new Triangle();
__Check((shape.Sides).ToString(), "3");

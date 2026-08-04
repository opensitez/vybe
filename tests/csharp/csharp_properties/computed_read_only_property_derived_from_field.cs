// vybe-test: csharp/csharp_properties/computed_read_only_property_derived_from_field
// origin: languages/csharp/tests/csharp/test_csharp_properties.rs

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

class Circle { public double Radius; public double Area => System.Math.PI * Radius * Radius; }
__P((System.Math.Round(new Circle{Radius=0}.Area)).ToString());
__Check("0");

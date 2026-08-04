// vybe-test: csharp/csharp_covariant_return_override/property_override_can_narrow_getter_return_type_covariantly
// origin: languages/csharp/tests/csharp/test_csharp_covariant_return_override.rs

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

class Shape { public virtual Shape CloneShape() { return new Shape(); } }
class Circle : Shape { public int Radius; public override Circle CloneShape() { return new Circle { Radius = Radius }; } }
Shape original = new Circle { Radius = 5 };
var copy = original.CloneShape();
__P((copy is Circle).ToString());
__Check("True");

// vybe-test: csharp/csharp_virtual_dispatch_semantics/virtual_property_getter_dispatches_to_derived_accessor
// origin: languages/csharp/tests/csharp/test_csharp_virtual_dispatch_semantics.rs

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

class Shape {
    public virtual int Sides { get { return 0; } }
}
class Triangle : Shape {
    public override int Sides { get { return 3; } }
}
Shape shape = new Triangle();
__P((shape.Sides).ToString());
__Check("3");

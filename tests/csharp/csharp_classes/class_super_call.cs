// vybe-test: csharp/csharp_classes/class_super_call
// origin: languages/csharp/tests/csharp/test_csharp_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Base {
    public virtual string Greet() { return "Hello"; }
}
class Derived : Base {
    public override string Greet() { return base.Greet() + " World"; }
}
var d = new Derived();
__Check((d.Greet()).ToString(), "Hello World");

// vybe-test: csharp/oop_advanced/virtual_base_call
// origin: languages/csharp/tests/csharp/test_oop_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Base {
    public virtual string Greet() { return "Hello"; }
}
class Child : Base {
    public override string Greet() { return base.Greet() + " World"; }
}
var c = new Child();
__Check((c.Greet()).ToString(), "Hello World");

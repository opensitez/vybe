// vybe-test: csharp/csharp_virtual_dispatch_semantics/calling_virtual_from_constructor_uses_current_type_override
// origin: languages/csharp/tests/csharp/test_csharp_virtual_dispatch_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Base {
    public Base() { __Check((Describe()).ToString(), ""); }
    public virtual string Describe() { return "base"; }
}
class Derived : Base {
    string label = "derived";
    public override string Describe() { return label; }
}
new Derived();

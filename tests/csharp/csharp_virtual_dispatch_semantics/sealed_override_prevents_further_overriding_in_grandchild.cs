// vybe-test: csharp/csharp_virtual_dispatch_semantics/sealed_override_prevents_further_overriding_in_grandchild
// origin: languages/csharp/tests/csharp/test_csharp_virtual_dispatch_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Base {
    public virtual string Tag() { return "base"; }
}
class Middle : Base {
    public sealed override string Tag() { return "middle"; }
}
class Leaf : Middle {
    public override string Tag() { return "leaf"; }
}
Base item = new Leaf();
__Check((item.Tag()).ToString(), "middle");

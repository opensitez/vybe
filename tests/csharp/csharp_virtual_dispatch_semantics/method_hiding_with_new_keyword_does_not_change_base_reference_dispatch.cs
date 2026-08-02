// vybe-test: csharp/csharp_virtual_dispatch_semantics/method_hiding_with_new_keyword_does_not_change_base_reference_dispatch
// origin: languages/csharp/tests/csharp/test_csharp_virtual_dispatch_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Base {
    public string Name() { return "base"; }
}
class Derived : Base {
    public new string Name() { return "derived"; }
}
Base reference = new Derived();
Derived concrete = new Derived();
__Check((reference.Name()).ToString(), "base");
__Check((concrete.Name()).ToString(), "derived");

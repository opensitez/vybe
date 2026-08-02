// vybe-test: csharp/csharp_virtual_dispatch_semantics/virtual_call_through_base_reference_uses_most_derived_override
// origin: languages/csharp/tests/csharp/test_csharp_virtual_dispatch_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Animal {
    public virtual string Speak() { return "..."; }
}
class Dog : Animal {
    public override string Speak() { return "woof"; }
}
Animal pet = new Dog();
__Check((pet.Speak()).ToString(), "woof");

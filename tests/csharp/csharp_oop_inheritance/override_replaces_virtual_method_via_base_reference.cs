// vybe-test: csharp/csharp_oop_inheritance/override_replaces_virtual_method_via_base_reference
// origin: languages/csharp/tests/csharp/test_csharp_oop_inheritance.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Animal { public virtual string Sound() => "..."; }
class Dog : Animal { public override string Sound() => "woof"; }
Animal a = new Dog();
__Check((a.Sound()).ToString(), "woof");

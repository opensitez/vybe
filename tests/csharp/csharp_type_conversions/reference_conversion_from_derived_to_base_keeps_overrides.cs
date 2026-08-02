// vybe-test: csharp/csharp_type_conversions/reference_conversion_from_derived_to_base_keeps_overrides
// origin: languages/csharp/tests/csharp/test_csharp_type_conversions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Animal { public virtual string Speak() { return "animal"; } } class Dog : Animal { public override string Speak() { return "woof"; } } Dog dog = new Dog(); Animal animal = dog; __Check((animal.Speak()).ToString(), "woof");

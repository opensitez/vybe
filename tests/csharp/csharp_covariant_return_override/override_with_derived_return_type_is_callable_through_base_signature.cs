// vybe-test: csharp/csharp_covariant_return_override/override_with_derived_return_type_is_callable_through_base_signature
// origin: languages/csharp/tests/csharp/test_csharp_covariant_return_override.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Animal { public string Name = "generic"; }
class Dog : Animal { }
class Shelter {
    public virtual Animal Adopt() { return new Animal(); }
}
class DogShelter : Shelter {
    public override Dog Adopt() { return new Dog { Name = "rex" }; }
}
Shelter place = new DogShelter();
__Check((place.Adopt().Name).ToString(), "rex");

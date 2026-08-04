// vybe-test: csharp/csharp_covariant_return_override/override_with_derived_return_type_is_callable_through_base_signature
// origin: languages/csharp/tests/csharp/test_csharp_covariant_return_override.rs

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

class Animal { public string Name = "generic"; }
class Dog : Animal { }
class Shelter {
    public virtual Animal Adopt() { return new Animal(); }
}
class DogShelter : Shelter {
    public override Dog Adopt() { return new Dog { Name = "rex" }; }
}
Shelter place = new DogShelter();
__P((place.Adopt().Name).ToString());
__Check("rex");

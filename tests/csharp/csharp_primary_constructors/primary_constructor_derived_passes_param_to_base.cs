// vybe-test: csharp/csharp_primary_constructors/primary_constructor_derived_passes_param_to_base
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Animal(string name) { public string Name => name; }
class Dog(string name, string breed) : Animal(name) { public string Breed => breed; }
var d = new Dog("Rex", "Lab");
__Check((d.Name).ToString(), "Rex"); __Check((d.Breed).ToString(), "Lab");

// vybe-test: csharp/csharp_oop/constructor_chaining_base
// origin: languages/csharp/tests/csharp/test_csharp_oop.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Animal {
    public string Name;
    public Animal(string name) { Name = name; }
}
class Dog : Animal {
    public string Breed;
    public Dog(string name, string breed) : base(name) {
        Breed = breed;
    }
}
var d = new Dog("Rex", "Lab");
__Check((d.Name).ToString(), "Rex");
__Check((d.Breed).ToString(), "Lab");

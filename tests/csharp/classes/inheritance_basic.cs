// vybe-test: csharp/classes/inheritance_basic
// origin: languages/csharp/tests/csharp/test_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Animal {
            string species;
            public Animal(string s) { this.species = s; }
            public string GetSpecies() { return this.species; }
        }
        class Dog : Animal {
            public Dog() : base("Canine") {}
        }
        var d = new Dog();
        __Check((d.GetSpecies()).ToString(), "Canine");

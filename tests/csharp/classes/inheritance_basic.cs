// vybe-test: csharp/classes/inheritance_basic
// origin: languages/csharp/tests/csharp/test_classes.rs

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

class Animal {
            string species;
            public Animal(string s) { this.species = s; }
            public string GetSpecies() { return this.species; }
        }
        class Dog : Animal {
            public Dog() : base("Canine") {}
        }
        var d = new Dog();
        __P((d.GetSpecies()).ToString());
__Check("Canine");

// vybe-test: csharp/classes/inheritance_override_method
// origin: languages/csharp/tests/csharp/test_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Animal {
            string name;
            public Animal(string n) { this.name = n; }
            public string Speak() { return this.name + " speaks"; }
        }
        class Dog : Animal {
            public Dog(string n) : base(n) {}
            public string Bark() { return this.name + " barks"; }
        }
        var d = new Dog("Rex");
        __Check((d.Speak()).ToString(), "Rex speaks");
        __Check((d.Bark()).ToString(), "Rex barks");

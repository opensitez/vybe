// vybe-test: csharp/classes/inheritance_override_method
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
            string name;
            public Animal(string n) { this.name = n; }
            public string Speak() { return this.name + " speaks"; }
        }
        class Dog : Animal {
            public Dog(string n) : base(n) {}
            public string Bark() { return this.name + " barks"; }
        }
        var d = new Dog("Rex");
        __P((d.Speak()).ToString());
        __P((d.Bark()).ToString());
__Check("Rex speaks\nRex barks");

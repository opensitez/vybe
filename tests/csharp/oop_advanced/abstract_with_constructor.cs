// vybe-test: csharp/oop_advanced/abstract_with_constructor
// origin: languages/csharp/tests/csharp/test_oop_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

abstract class Animal {
    protected string name;
    public Animal(string n) { name = n; }
    public abstract string Sound();
    public string Greet() { return name + " says " + Sound(); }
}
class Dog : Animal {
    public Dog(string n) : base(n) { }
    public override string Sound() { return "Woof"; }
}
var d = new Dog("Rex");
__Check((d.Greet()).ToString(), "Rex says Woof");

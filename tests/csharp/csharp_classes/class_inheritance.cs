// vybe-test: csharp/csharp_classes/class_inheritance
// origin: languages/csharp/tests/csharp/test_csharp_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Animal {
    public string Name;
    public Animal(string name) { Name = name; }
    public virtual string Speak() { return Name + " speaks"; }
}
class Dog : Animal {
    public Dog(string name) : base(name) {}
    public override string Speak() { return Name + " barks"; }
}
var d = new Dog("Rex");
__Check((d.Speak()).ToString(), "Rex barks");

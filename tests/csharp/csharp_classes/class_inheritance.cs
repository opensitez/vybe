// vybe-test: csharp/csharp_classes/class_inheritance
// origin: languages/csharp/tests/csharp/test_csharp_classes.rs

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
    public string Name;
    public Animal(string name) { Name = name; }
    public virtual string Speak() { return Name + " speaks"; }
}
class Dog : Animal {
    public Dog(string name) : base(name) {}
    public override string Speak() { return Name + " barks"; }
}
var d = new Dog("Rex");
__P((d.Speak()).ToString());
__Check("Rex barks");

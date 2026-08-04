// vybe-test: csharp/oop_advanced/polymorphic_list
// origin: languages/csharp/tests/csharp/test_oop_advanced.rs

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
    public virtual string Speak() { return "..."; }
}
class Dog : Animal {
    public override string Speak() { return "Woof"; }
}
class Cat : Animal {
    public override string Speak() { return "Meow"; }
}
var animals = new List<Animal> { new Dog(), new Cat(), new Dog() };
foreach (var a in animals) {
    __P((a.Speak()).ToString());
}
__Check("Woof\nMeow\nWoof");

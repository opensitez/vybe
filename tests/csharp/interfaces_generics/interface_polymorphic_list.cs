// vybe-test: csharp/interfaces_generics/interface_polymorphic_list
// origin: languages/csharp/tests/csharp/test_interfaces_generics.rs

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

interface IAnimal {
    string Speak();
}
class Dog : IAnimal {
    public string Speak() { return "Woof"; }
}
class Cat : IAnimal {
    public string Speak() { return "Meow"; }
}
var animals = new List<IAnimal> { new Dog(), new Cat(), new Dog() };
foreach (var a in animals) __P((a.Speak()).ToString());
__Check("Woof\nMeow\nWoof");

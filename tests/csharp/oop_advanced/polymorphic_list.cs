// vybe-test: csharp/oop_advanced/polymorphic_list
// origin: languages/csharp/tests/csharp/test_oop_advanced.rs

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
    Console.WriteLine(a.Speak());
}

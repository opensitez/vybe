// vybe-test: csharp/interfaces_generics/interface_polymorphic_list
// origin: languages/csharp/tests/csharp/test_interfaces_generics.rs

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
foreach (var a in animals) Console.WriteLine(a.Speak());

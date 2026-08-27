// vybe-test: csharp/interfaces_generics/interface_polymorphic_list
// origin: languages/csharp/tests/csharp/test_interfaces_generics.rs

using static __Harness;

var animals = new List<IAnimal> { new Dog(), new Cat(), new Dog() }
;
foreach (var a in animals) __P((a.Speak()).ToString());
__Check("Woof\nMeow\nWoof");

interface IAnimal {
    string Speak();
}

class Dog : IAnimal {
    public string Speak() { return "Woof"; }
}

class Cat : IAnimal {
    public string Speak() { return "Meow"; }
}

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}

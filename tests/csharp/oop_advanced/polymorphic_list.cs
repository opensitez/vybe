// vybe-test: csharp/oop_advanced/polymorphic_list
// origin: languages/csharp/tests/csharp/test_oop_advanced.rs

using static __Harness;

var animals = new List<Animal> { new Dog(), new Cat(), new Dog() }
;
foreach (var a in animals) {
    __P((a.Speak()).ToString());
}
__Check("Woof\nMeow\nWoof");

class Animal {
    public virtual string Speak() { return "..."; }
}

class Dog : Animal {
    public override string Speak() { return "Woof"; }
}

class Cat : Animal {
    public override string Speak() { return "Meow"; }
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

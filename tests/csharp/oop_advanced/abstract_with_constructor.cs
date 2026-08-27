// vybe-test: csharp/oop_advanced/abstract_with_constructor
// origin: languages/csharp/tests/csharp/test_oop_advanced.rs

using static __Harness;

var d = new Dog("Rex");
__P((d.Greet()).ToString());
__Check("Rex says Woof");

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

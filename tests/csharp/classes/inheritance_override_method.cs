// vybe-test: csharp/classes/inheritance_override_method
// origin: languages/csharp/tests/csharp/test_classes.rs

using static __Harness;

var d = new Dog("Rex");
__P((d.Speak()).ToString());
__P((d.Bark()).ToString());
__Check("Rex speaks\nRex barks");

class Animal {
            protected string name;
            public Animal(string n) { this.name = n; }
            public string Speak() { return this.name + " speaks"; }
        }

class Dog : Animal {
            public Dog(string n) : base(n) {}
            public string Bark() { return this.name + " barks"; }
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

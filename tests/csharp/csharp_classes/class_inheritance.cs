// vybe-test: csharp/csharp_classes/class_inheritance
// origin: languages/csharp/tests/csharp/test_csharp_classes.rs

using static __Harness;

var d = new Dog("Rex");
__P((d.Speak()).ToString());
__Check("Rex barks");

class Animal {
    public string Name;
    public Animal(string name) { Name = name; }
    public virtual string Speak() { return Name + " speaks"; }
}

class Dog : Animal {
    public Dog(string name) : base(name) {}
    public override string Speak() { return Name + " barks"; }
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

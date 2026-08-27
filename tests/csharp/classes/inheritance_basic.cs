// vybe-test: csharp/classes/inheritance_basic
// origin: languages/csharp/tests/csharp/test_classes.rs

using static __Harness;

var d = new Dog();
__P((d.GetSpecies()).ToString());
__Check("Canine");

class Animal {
            string species;
            public Animal(string s) { this.species = s; }
            public string GetSpecies() { return this.species; }
        }

class Dog : Animal {
            public Dog() : base("Canine") {}
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

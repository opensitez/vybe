// vybe-test: csharp/csharp_oop/constructor_chaining_base
// origin: languages/csharp/tests/csharp/test_csharp_oop.rs

using static __Harness;

var d = new Dog("Rex", "Lab");
__P((d.Name).ToString());
__P((d.Breed).ToString());
__Check("Rex\nLab");

class Animal {
    public string Name;
    public Animal(string name) { Name = name; }
}

class Dog : Animal {
    public string Breed;
    public Dog(string name, string breed) : base(name) {
        Breed = breed;
    }
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

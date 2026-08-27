// vybe-test: csharp/csharp_primary_constructors/primary_constructor_derived_passes_param_to_base
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

using static __Harness;

var d = new Dog("Rex", "Lab");
__P((d.Name).ToString());
__P((d.Breed).ToString());
__Check("Rex\nLab");

class Animal(string name) { public string Name => name; }

class Dog(string name, string breed) : Animal(name) { public string Breed => breed; }

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

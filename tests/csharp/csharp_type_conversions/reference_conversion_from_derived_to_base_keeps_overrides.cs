// vybe-test: csharp/csharp_type_conversions/reference_conversion_from_derived_to_base_keeps_overrides
// origin: languages/csharp/tests/csharp/test_csharp_type_conversions.rs

using static __Harness;

Dog dog = new Dog();
Animal animal = dog;
__P((animal.Speak()).ToString());
__Check("woof");

class Animal { public virtual string Speak() { return "animal"; } }

class Dog : Animal { public override string Speak() { return "woof"; } }

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

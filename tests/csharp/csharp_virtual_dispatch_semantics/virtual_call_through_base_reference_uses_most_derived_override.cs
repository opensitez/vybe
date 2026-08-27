// vybe-test: csharp/csharp_virtual_dispatch_semantics/virtual_call_through_base_reference_uses_most_derived_override
// origin: languages/csharp/tests/csharp/test_csharp_virtual_dispatch_semantics.rs

using static __Harness;

Animal pet = new Dog();
__P((pet.Speak()).ToString());
__Check("woof");

class Animal {
    public virtual string Speak() { return "..."; }
}

class Dog : Animal {
    public override string Speak() { return "woof"; }
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

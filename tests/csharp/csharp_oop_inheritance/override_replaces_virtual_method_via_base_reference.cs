// vybe-test: csharp/csharp_oop_inheritance/override_replaces_virtual_method_via_base_reference
// origin: languages/csharp/tests/csharp/test_csharp_oop_inheritance.rs

using static __Harness;

Animal a = new Dog();
__P((a.Sound()).ToString());
__Check("woof");

class Animal { public virtual string Sound() => "..."; }

class Dog : Animal { public override string Sound() => "woof"; }

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

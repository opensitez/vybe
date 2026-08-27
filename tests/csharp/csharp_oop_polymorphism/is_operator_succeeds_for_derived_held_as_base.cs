// vybe-test: csharp/csharp_oop_polymorphism/is_operator_succeeds_for_derived_held_as_base
// origin: languages/csharp/tests/csharp/test_csharp_oop_polymorphism.rs

using static __Harness;

Animal a=new Dog();
__P((a is Dog).ToString());
__P((a is Animal).ToString());
__Check("True\nTrue");

class Animal{}

class Dog:Animal{}

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

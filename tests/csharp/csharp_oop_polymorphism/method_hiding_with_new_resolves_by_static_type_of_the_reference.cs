// vybe-test: csharp/csharp_oop_polymorphism/method_hiding_with_new_resolves_by_static_type_of_the_reference
// origin: languages/csharp/tests/csharp/test_csharp_oop_polymorphism.rs

using static __Harness;

Derived d=new Derived();
Base b=d;
__P((d.Speak()).ToString());
__P((b.Speak()).ToString());
__Check("hidden\nbase");

class Base{public virtual string Speak()=>"base";}

class Derived:Base{public new string Speak()=>"hidden";}

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

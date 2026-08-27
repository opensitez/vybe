// vybe-test: csharp/csharp_oop_polymorphism/virtual_method_dispatches_to_most_derived_override
// origin: languages/csharp/tests/csharp/test_csharp_oop_polymorphism.rs

using static __Harness;

Base obj=new Derived();
__P((obj.Speak()).ToString());
__Check("derived");

class Base{public virtual string Speak()=>"base";}

class Derived:Base{public override string Speak()=>"derived";}

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

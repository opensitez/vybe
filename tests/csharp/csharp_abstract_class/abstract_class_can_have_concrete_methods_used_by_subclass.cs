// vybe-test: csharp/csharp_abstract_class/abstract_class_can_have_concrete_methods_used_by_subclass
// origin: languages/csharp/tests/csharp/test_csharp_abstract_class.rs

using static __Harness;

__P((new Cat().Speak()).ToString());
__Check("I say meow");

abstract class Animal{
    public abstract string Sound();
    public string Speak()=>$"I say {Sound()}";
}

class Cat:Animal{public override string Sound()=>"meow";}

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

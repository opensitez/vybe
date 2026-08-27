// vybe-test: csharp/csharp_design_patterns/factory_method_creates_correct_concrete_type
// origin: languages/csharp/tests/csharp/test_csharp_design_patterns.rs

using static __Harness;

Animal Create(string kind)=>kind=="dog"?(Animal)new Dog():new Cat();
__P((Create("dog").Sound()).ToString());
__P((Create("cat").Sound()).ToString());
__Check("woof\nmeow");

abstract class Animal{public abstract string Sound();}

class Dog:Animal{public override string Sound()=>"woof";}

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

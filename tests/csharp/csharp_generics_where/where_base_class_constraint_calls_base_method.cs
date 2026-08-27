// vybe-test: csharp/csharp_generics_where/where_base_class_constraint_calls_base_method
// origin: languages/csharp/tests/csharp/test_csharp_generics_where.rs

using static __Harness;

string Hear<T>(T a) where T:Animal=>a.Sound();
__P((Hear(new Dog())).ToString());
__Check("woof");

abstract class Animal{public abstract string Sound();}

class Dog:Animal{public override string Sound()=>"woof";}

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

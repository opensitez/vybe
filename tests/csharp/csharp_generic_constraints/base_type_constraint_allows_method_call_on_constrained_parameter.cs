// vybe-test: csharp/csharp_generic_constraints/base_type_constraint_allows_method_call_on_constrained_parameter
// origin: languages/csharp/tests/csharp/test_csharp_generic_constraints.rs

using static __Harness;

string Speak<T>(T t) where T : Animal => t.Sound();
__P((Speak(new Dog())).ToString());
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

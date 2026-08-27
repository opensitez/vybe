// vybe-test: csharp/csharp_abstract_sealed/sealed_method_override_stops_further_overriding_in_chain
// origin: languages/csharp/tests/csharp/test_csharp_abstract_sealed.rs

using static __Harness;

A obj = new C();
__P((obj.Name()).ToString());
__Check("B");

class A { public virtual string Name() => "A"; }

class B : A { public sealed override string Name() => "B"; }

class C : B { }

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

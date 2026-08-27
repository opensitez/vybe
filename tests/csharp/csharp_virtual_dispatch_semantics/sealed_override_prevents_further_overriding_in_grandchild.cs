// vybe-test: csharp/csharp_virtual_dispatch_semantics/sealed_override_prevents_further_overriding_in_grandchild
// origin: languages/csharp/tests/csharp/test_csharp_virtual_dispatch_semantics.rs

using static __Harness;

Base b = new Child();
b.Work();
__Check("child");

class Base {
    public virtual void Work() => __P("base");
}
class Child : Base {
    public sealed override void Work() => __P("child");
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

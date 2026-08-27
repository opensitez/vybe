// vybe-test: csharp/csharp_virtual_dispatch_semantics/calling_virtual_from_constructor_uses_current_type_override
// origin: languages/csharp/tests/csharp/test_csharp_virtual_dispatch_semantics.rs

using static __Harness;

var d = new Derived();
__Check("derived");

class Base {
    public Base() { Init(); }
    public virtual void Init() => __P("base");
}
class Derived : Base {
    public override void Init() => __P("derived");
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

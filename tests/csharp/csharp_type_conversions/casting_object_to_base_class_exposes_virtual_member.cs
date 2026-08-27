// vybe-test: csharp/csharp_type_conversions/casting_object_to_base_class_exposes_virtual_member
// origin: languages/csharp/tests/csharp/test_csharp_type_conversions.rs

using static __Harness;

object item = new Child();
__P((((Base)item).Name()).ToString());
__Check("child");

class Base { public virtual string Name() { return "base"; } }

class Child : Base { public override string Name() { return "child"; } }

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

// vybe-test: csharp/csharp_nested_type_member_access/nested_class_can_read_outer_private_instance_field
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_member_access.rs

using static __Harness;

__P((new Outer().ViaInner()).ToString());
__Check("8");

class Outer {
    int secret = 8;
    class Inner {
        Outer parent;
        public Inner(Outer parent) { this.parent = parent; }
        public int Read() { return parent.secret; }
    }
    public int ViaInner() { return new Inner(this).Read(); }
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

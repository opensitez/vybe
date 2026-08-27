// vybe-test: csharp/csharp_nested_type_member_access/nested_class_can_invoke_outer_private_instance_method
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_member_access.rs

using static __Harness;

__P((new Outer().ViaInner(5)).ToString());
__Check("10");

class Outer {
    int Twice(int n) { return n * 2; }
    class Inner {
        Outer parent;
        public Inner(Outer parent) { this.parent = parent; }
        public int Run(int n) { return parent.Twice(n); }
    }
    public int ViaInner(int n) { return new Inner(this).Run(n); }
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

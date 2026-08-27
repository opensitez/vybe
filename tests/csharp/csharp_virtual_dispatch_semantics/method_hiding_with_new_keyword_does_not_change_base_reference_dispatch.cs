// vybe-test: csharp/csharp_virtual_dispatch_semantics/method_hiding_with_new_keyword_does_not_change_base_reference_dispatch
// origin: languages/csharp/tests/csharp/test_csharp_virtual_dispatch_semantics.rs

using static __Harness;

Base reference = new Derived();
Derived concrete = new Derived();
__P((reference.Name()).ToString());
__P((concrete.Name()).ToString());
__Check("base\nderived");

class Base {
    public string Name() { return "base"; }
}

class Derived : Base {
    public new string Name() { return "derived"; }
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

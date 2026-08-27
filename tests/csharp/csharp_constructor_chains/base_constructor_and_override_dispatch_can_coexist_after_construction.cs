// vybe-test: csharp/csharp_constructor_chains/base_constructor_and_override_dispatch_can_coexist_after_construction
// origin: languages/csharp/tests/csharp/test_csharp_constructor_chains.rs

using static __Harness;

__P((new Child().Read()).ToString());
__Check("xy");

class Base { protected string prefix; public Base(string prefix) { this.prefix = prefix; } public virtual string Read() { return prefix; } }

class Child : Base { public Child() : base("x") { } public override string Read() { return prefix + "y"; } }

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

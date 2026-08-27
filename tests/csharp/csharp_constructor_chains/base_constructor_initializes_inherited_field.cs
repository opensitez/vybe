// vybe-test: csharp/csharp_constructor_chains/base_constructor_initializes_inherited_field
// origin: languages/csharp/tests/csharp/test_csharp_constructor_chains.rs

using static __Harness;

__P((new Child().Name()).ToString());
__Check("root");

class Base { protected string name; public Base(string name) { this.name = name; } public string Name() { return name; } }

class Child : Base { public Child() : base("root") { } }

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

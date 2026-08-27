// vybe-test: csharp/csharp_constructor_chains/base_and_derived_constructors_run_in_order
// origin: languages/csharp/tests/csharp/test_csharp_constructor_chains.rs

using static __Harness;

new Child();
__Check("base\nchild");

class Base { public Base() { __P(("base").ToString()); } }

class Child : Base { public Child() { __P(("child").ToString()); } }

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

// vybe-test: csharp/csharp_constructor_chains/derived_constructor_can_pass_argument_from_expression_to_base
// origin: languages/csharp/tests/csharp/test_csharp_constructor_chains.rs

using static __Harness;

__P((new Child(4).Read()).ToString());
__Check("5");

class Base { int value; public Base(int value) { this.value = value; } public int Read() { return value; } }

class Child : Base { public Child(int value) : base(value + 1) { } }

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

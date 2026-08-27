// vybe-test: csharp/csharp_constructor_chains/constructor_chain_can_set_multiple_fields_from_single_input
// origin: languages/csharp/tests/csharp/test_csharp_constructor_chains.rs

using static __Harness;

__P((new Box("a").Read()).ToString());
__Check("a:A");

class Box { string left; string right; public Box(string value) : this(value, value.ToUpper()) { } public Box(string left, string right) { this.left = left; this.right = right; } public string Read() { return left + ":" + right; } }

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

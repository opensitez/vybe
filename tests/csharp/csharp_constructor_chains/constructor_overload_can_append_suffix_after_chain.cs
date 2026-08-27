// vybe-test: csharp/csharp_constructor_chains/constructor_overload_can_append_suffix_after_chain
// origin: languages/csharp/tests/csharp/test_csharp_constructor_chains.rs

using static __Harness;

__P((new Box("a", "b").Read()).ToString());
__Check("ab");

class Box { string name; public Box(string name) { this.name = name; } public Box(string name, string suffix) : this(name) { this.name += suffix; } public string Read() { return name; } }

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

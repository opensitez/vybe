// vybe-test: csharp/csharp_constructor_chains/this_constructor_chain_reuses_primary_overload
// origin: languages/csharp/tests/csharp/test_csharp_constructor_chains.rs

using static __Harness;

__P((new Box().Read()).ToString());
__Check("9");

class Box { int value; public Box() : this(9) { } public Box(int value) { this.value = value; } public int Read() { return value; } }

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

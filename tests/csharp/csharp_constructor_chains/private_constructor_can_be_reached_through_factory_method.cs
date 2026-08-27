// vybe-test: csharp/csharp_constructor_chains/private_constructor_can_be_reached_through_factory_method
// origin: languages/csharp/tests/csharp/test_csharp_constructor_chains.rs

using static __Harness;

__P((Box.Create().Read()).ToString());
__Check("made");

class Box { string name; private Box(string name) { this.name = name; } public static Box Create() { return new Box("made"); } public string Read() { return name; } }

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

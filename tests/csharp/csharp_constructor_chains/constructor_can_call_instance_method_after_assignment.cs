// vybe-test: csharp/csharp_constructor_chains/constructor_can_call_instance_method_after_assignment
// origin: languages/csharp/tests/csharp/test_csharp_constructor_chains.rs

using static __Harness;

new Box(8);
__Check("8");

class Box { int value; public Box(int value) { this.value = value; __P((Read()).ToString()); } public int Read() { return value; } }

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

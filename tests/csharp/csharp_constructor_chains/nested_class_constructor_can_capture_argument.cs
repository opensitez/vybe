// vybe-test: csharp/csharp_constructor_chains/nested_class_constructor_can_capture_argument
// origin: languages/csharp/tests/csharp/test_csharp_constructor_chains.rs

using static __Harness;

__P((new Outer.Inner("inner").Read()).ToString());
__Check("inner");

class Outer { public class Inner { string name; public Inner(string name) { this.name = name; } public string Read() { return name; } } }

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

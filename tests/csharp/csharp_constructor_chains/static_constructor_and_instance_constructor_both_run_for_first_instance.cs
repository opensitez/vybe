// vybe-test: csharp/csharp_constructor_chains/static_constructor_and_instance_constructor_both_run_for_first_instance
// origin: languages/csharp/tests/csharp/test_csharp_constructor_chains.rs

using static __Harness;

new Box();
__Check("static\ninstance");

class Box { static Box() { __P(("static").ToString()); } public Box() { __P(("instance").ToString()); } }

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

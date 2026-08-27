// vybe-test: csharp/csharp_queue_stack_surface/queue_stack_surface_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_surface.rs

using static __Harness;

// queue_stack_surface
double seed = 32;
__P(((seed + 0.5 - 0.5) == seed).ToString());
__Check("True");

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

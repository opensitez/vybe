// vybe-test: csharp/csharp_queue_stack_surface/queue_stack_surface_arithmetic_increment
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_surface.rs

using static __Harness;

// queue_stack_surface
int seed = 32;
__P((seed + 1 > seed).ToString());
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

// vybe-test: csharp/csharp_environment_system/gc_get_total_memory_returns_positive_long
// origin: languages/csharp/tests/csharp/test_csharp_environment_system.rs

using static __Harness;

__P((System.GC.GetTotalMemory(false)>0).ToString());
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

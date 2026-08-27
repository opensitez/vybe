// vybe-test: csharp/csharp_volatile_thread_memory/volatile_static_bool_flag_count
// origin: languages/csharp/tests/csharp/test_csharp_volatile_thread_memory.rs

using static __Harness;

FlagBox.Done = true;
__P((FlagBox.Done ? 1 : 0).ToString());
__Check("1");

class FlagBox {
    public static volatile bool Done = false;
}

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

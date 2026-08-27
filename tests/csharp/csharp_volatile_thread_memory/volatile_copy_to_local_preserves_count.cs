// vybe-test: csharp/csharp_volatile_thread_memory/volatile_copy_to_local_preserves_count
// origin: languages/csharp/tests/csharp/test_csharp_volatile_thread_memory.rs

using static __Harness;

var box = new FlagBox();
int local = box.Value;
__P((local).ToString());
__Check("8");

class FlagBox {
    public volatile int Value = 8;
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

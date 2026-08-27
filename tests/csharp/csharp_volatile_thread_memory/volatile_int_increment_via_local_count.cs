// vybe-test: csharp/csharp_volatile_thread_memory/volatile_int_increment_via_local_count
// origin: languages/csharp/tests/csharp/test_csharp_volatile_thread_memory.rs

using static __Harness;

var box = new FlagBox();
int snapshot = box.Value;
box.Value = snapshot + 2;
__P((box.Value).ToString());
__Check("3");

class FlagBox {
    public volatile int Value = 1;
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

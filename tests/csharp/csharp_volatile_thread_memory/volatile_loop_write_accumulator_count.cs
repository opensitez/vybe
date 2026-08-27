// vybe-test: csharp/csharp_volatile_thread_memory/volatile_loop_write_accumulator_count
// origin: languages/csharp/tests/csharp/test_csharp_volatile_thread_memory.rs

using static __Harness;

var box = new FlagBox();
for (int i = 1; i <= 4; i++) box.Value += i;
__P((box.Value).ToString());
__Check("10");

class FlagBox {
    public volatile int Value = 0;
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

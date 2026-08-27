// vybe-test: csharp/csharp_volatile_thread_memory/volatile_read_while_loop_count
// origin: languages/csharp/tests/csharp/test_csharp_volatile_thread_memory.rs

using static __Harness;

var box = new FlagBox();
int count = 0;
while (box.Value > 0) {
    count++;
    box.Value--;
}
__P((count).ToString());
__Check("3");

class FlagBox {
    public volatile int Value = 3;
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

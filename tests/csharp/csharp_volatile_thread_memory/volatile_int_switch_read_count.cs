// vybe-test: csharp/csharp_volatile_thread_memory/volatile_int_switch_read_count
// origin: languages/csharp/tests/csharp/test_csharp_volatile_thread_memory.rs

using static __Harness;

var box = new FlagBox();
int count = 0;
switch (box.Value) {
    case 3: count = 30; break;
    default: count = 0; break;
}
__P((count).ToString());
__Check("30");

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

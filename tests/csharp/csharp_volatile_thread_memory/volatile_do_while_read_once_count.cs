// vybe-test: csharp/csharp_volatile_thread_memory/volatile_do_while_read_once_count
// origin: languages/csharp/tests/csharp/test_csharp_volatile_thread_memory.rs

using static __Harness;

var box = new FlagBox();
int count = 0;
do {
    count += box.Value;
    box.Value = 0;
}
while (box.Value > 0);
__P((count).ToString());
__Check("1");

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

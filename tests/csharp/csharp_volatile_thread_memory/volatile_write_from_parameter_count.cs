// vybe-test: csharp/csharp_volatile_thread_memory/volatile_write_from_parameter_count
// origin: languages/csharp/tests/csharp/test_csharp_volatile_thread_memory.rs

using static __Harness;

var box = new FlagBox();
box.Set(22);
__P((box.Value).ToString());
__Check("22");

class FlagBox {
    public volatile int Value = 0;
    public void Set(int n) { Value = n; }
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

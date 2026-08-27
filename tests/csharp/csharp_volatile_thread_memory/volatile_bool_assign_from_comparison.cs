// vybe-test: csharp/csharp_volatile_thread_memory/volatile_bool_assign_from_comparison
// origin: languages/csharp/tests/csharp/test_csharp_volatile_thread_memory.rs

using static __Harness;

var box = new FlagBox();
box.Ready = 5 > 3;
__P((box.Ready ? 1 : 0).ToString());
__Check("1");

class FlagBox {
    public volatile bool Ready = false;
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

// vybe-test: csharp/csharp_volatile_thread_memory/volatile_bool_or_expression_count
// origin: languages/csharp/tests/csharp/test_csharp_volatile_thread_memory.rs

using static __Harness;

var box = new FlagBox();
__P(((box.A || box.B) ? 1 : 0).ToString());
__Check("1");

class FlagBox {
    public volatile bool A = false;
    public volatile bool B = true;
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

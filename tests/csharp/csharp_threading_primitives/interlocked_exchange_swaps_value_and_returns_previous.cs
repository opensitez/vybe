// vybe-test: csharp/csharp_threading_primitives/interlocked_exchange_swaps_value_and_returns_previous
// origin: languages/csharp/tests/csharp/test_csharp_threading_primitives.rs

using static __Harness;

int slot = 1;
__P((System.Threading.Interlocked.Exchange(ref slot, 9)).ToString());
__P((slot).ToString());
__Check("1\n9");

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

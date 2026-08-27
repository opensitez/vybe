// vybe-test: csharp/csharp_interlocked_atomic/interlocked_compare_exchange_idempotent_same
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic.rs

using static __Harness;

int slot = 3;
var p1 = System.Threading.Interlocked.CompareExchange(ref slot, 8, 3);
var p2 = System.Threading.Interlocked.CompareExchange(ref slot, 8, 3);
__P((p1 + p2).ToString());
__P((slot).ToString());
__Check("11\n8");

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

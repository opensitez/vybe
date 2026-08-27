// vybe-test: csharp/csharp_interlocked_atomic/interlocked_exchange_overwrites_existing
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic.rs

using static __Harness;

int slot = 77;
__P((System.Threading.Interlocked.Exchange(ref slot, 3)).ToString());
__P((slot).ToString());
__Check("77\n3");

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

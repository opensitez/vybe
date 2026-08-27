// vybe-test: csharp/csharp_concurrency_sync/interlocked_compare_exchange_sets_only_when_expected
// origin: languages/csharp/tests/csharp/test_csharp_concurrency_sync.rs

using static __Harness;

int val=0;
int original=System.Threading.Interlocked.CompareExchange(ref val,99,0);
__P((original).ToString());
__P((val).ToString());
__Check("0\n99");

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

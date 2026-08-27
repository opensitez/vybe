// vybe-test: csharp/csharp_interlocked_atomic/interlocked_add_positive_delta
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic.rs

using static __Harness;

int total = 10;
__P((System.Threading.Interlocked.Add(ref total, 4)).ToString());
__P((total).ToString());
__Check("14\n14");

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

// vybe-test: csharp/csharp_interlocked_atomic/interlocked_increment_twice_count
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic.rs

using static __Harness;

int counter = 0;
System.Threading.Interlocked.Increment(ref counter);
__P((System.Threading.Interlocked.Increment(ref counter)).ToString());
__P((counter).ToString());
__Check("2\n2");

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

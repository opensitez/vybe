// vybe-test: csharp/csharp_interlocked_atomic/interlocked_add_multiple_steps_sum
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic.rs

using static __Harness;

int total = 0;
System.Threading.Interlocked.Add(ref total, 2);
System.Threading.Interlocked.Add(ref total, 3);
System.Threading.Interlocked.Add(ref total, 5);
__P((total).ToString());
__Check("10");

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

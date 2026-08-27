// vybe-test: csharp/csharp_interlocked_atomic/interlocked_add_one_equals_increment
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic.rs

using static __Harness;

int a = 6;
int b = 6;
System.Threading.Interlocked.Increment(ref a);
System.Threading.Interlocked.Add(ref b, 1);
__P((a + b).ToString());
__Check("14");

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

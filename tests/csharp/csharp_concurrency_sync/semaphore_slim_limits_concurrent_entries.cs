// vybe-test: csharp/csharp_concurrency_sync/semaphore_slim_limits_concurrent_entries
// origin: languages/csharp/tests/csharp/test_csharp_concurrency_sync.rs

using static __Harness;

var sem=new System.Threading.SemaphoreSlim(1,1);
sem.Wait();
__P((sem.CurrentCount).ToString());
sem.Release();
__P((sem.CurrentCount).ToString());
__Check("0\n1");

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

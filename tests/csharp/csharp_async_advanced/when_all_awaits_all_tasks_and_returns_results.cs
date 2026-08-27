// vybe-test: csharp/csharp_async_advanced/when_all_awaits_all_tasks_and_returns_results
// origin: languages/csharp/tests/csharp/test_csharp_async_advanced.rs

using static __Harness;

async System.Threading.Tasks.Task<int> N(int v){
    await System.Threading.Tasks.Task.Delay(0);return v;
}
int[] results=await System.Threading.Tasks.Task.WhenAll(N(1),N(2),N(3));
__P((results.Sum()).ToString());
__Check("6");

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

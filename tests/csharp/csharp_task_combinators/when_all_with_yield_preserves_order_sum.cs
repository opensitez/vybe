// vybe-test: csharp/csharp_task_combinators/when_all_with_yield_preserves_order_sum
// origin: languages/csharp/tests/csharp/test_csharp_task_combinators.rs

using static __Harness;

async System.Threading.Tasks.Task<int> Val(int n) {
    await System.Threading.Tasks.Task.Yield();
    return n;
}
async System.Threading.Tasks.Task Run() {
    var results = await System.Threading.Tasks.Task.WhenAll(Val(10), Val(20));
    __P((results[0] + results[1]).ToString());
}
Run().Wait();
__Check("30");

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

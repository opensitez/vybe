// vybe-test: csharp/csharp_task_combinators/when_all_three_tasks_sum_count
// origin: languages/csharp/tests/csharp/test_csharp_task_combinators.rs

using static __Harness;

async System.Threading.Tasks.Task<int> N(int v) { return v; }
async System.Threading.Tasks.Task Run() {
    var results = await System.Threading.Tasks.Task.WhenAll(N(1), N(2), N(3));
    int count = 0;
    foreach (var x in results) count += x;
    __P((count).ToString());
}
Run().Wait();
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

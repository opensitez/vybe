// vybe-test: csharp/csharp_value_task/value_task_three_way_sum_count
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

using static __Harness;

async System.Threading.Tasks.ValueTask<int> N(int v) { return v; }
async System.Threading.Tasks.Task Run() {
    int count = await N(4) + await N(5) + await N(6);
    __P((count).ToString());
}
Run().Wait();
__Check("15");

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

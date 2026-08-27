// vybe-test: csharp/csharp_value_task/value_task_delay_via_as_task_count
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

using static __Harness;

async System.Threading.Tasks.ValueTask<int> Delayed() {
    await System.Threading.Tasks.Task.Delay(0).ConfigureAwait(false);
    return 2;
}
async System.Threading.Tasks.Task Run() {
    var task = Delayed().AsTask();
    __P((await task).ToString());
}
Run().Wait();
__Check("2");

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

// vybe-test: csharp/csharp_value_task/value_task_completed_task_await_count
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

using static __Harness;

async System.Threading.Tasks.Task Run() {
    await System.Threading.Tasks.ValueTask.CompletedTask;
    __P((1).ToString());
}
Run().Wait();
__Check("1");

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

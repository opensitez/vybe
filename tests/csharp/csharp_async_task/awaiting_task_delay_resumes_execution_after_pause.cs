// vybe-test: csharp/csharp_async_task/awaiting_task_delay_resumes_execution_after_pause
// origin: languages/csharp/tests/csharp/test_csharp_async_task.rs

using static __Harness;

async System.Threading.Tasks.Task Run() {
    __P(("before").ToString());
    await System.Threading.Tasks.Task.Delay(1);
    __P(("after").ToString());
}
Run().Wait();
__Check("before\nafter");

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

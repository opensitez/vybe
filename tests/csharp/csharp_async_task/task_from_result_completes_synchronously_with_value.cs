// vybe-test: csharp/csharp_async_task/task_from_result_completes_synchronously_with_value
// origin: languages/csharp/tests/csharp/test_csharp_async_task.rs

using static __Harness;

async System.Threading.Tasks.Task Run() {
    var t = System.Threading.Tasks.Task.FromResult(42);
    __P((await t).ToString());
}
Run().Wait();
__Check("42");

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

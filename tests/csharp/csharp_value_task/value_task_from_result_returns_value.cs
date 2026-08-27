// vybe-test: csharp/csharp_value_task/value_task_from_result_returns_value
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

using static __Harness;

async System.Threading.Tasks.Task Run() {
    var vt = System.Threading.Tasks.ValueTask.FromResult(17);
    __P((await vt).ToString());
}
Run().Wait();
__Check("17");

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

// vybe-test: csharp/csharp_async_task/task_when_all_waits_for_multiple_tasks_to_complete
// origin: languages/csharp/tests/csharp/test_csharp_async_task.rs

using static __Harness;

async System.Threading.Tasks.Task<int> Val(int n) {
    await System.Threading.Tasks.Task.Yield();
    return n;
}
var results = System.Threading.Tasks.Task.WhenAll(Val(1), Val(2), Val(3)).Result;
int sum = 0;
foreach (var r in results) sum += r;
__P((sum).ToString());
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

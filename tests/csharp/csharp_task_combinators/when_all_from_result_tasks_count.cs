// vybe-test: csharp/csharp_task_combinators/when_all_from_result_tasks_count
// origin: languages/csharp/tests/csharp/test_csharp_task_combinators.rs

using static __Harness;

async System.Threading.Tasks.Task Run() {
    var results = await System.Threading.Tasks.Task.WhenAll(
        System.Threading.Tasks.Task.FromResult(7),
        System.Threading.Tasks.Task.FromResult(8)
    );
    __P((results[0] + results[1]).ToString());
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

// vybe-test: csharp/csharp_task_combinators/when_any_two_identical_fast_tasks_count
// origin: languages/csharp/tests/csharp/test_csharp_task_combinators.rs

using static __Harness;

async System.Threading.Tasks.Task<int> Fast() { return 4; }
async System.Threading.Tasks.Task Run() {
    var winner = await System.Threading.Tasks.Task.WhenAny(Fast(), Fast());
    __P((winner.Result).ToString());
}
Run().Wait();
__Check("4");

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

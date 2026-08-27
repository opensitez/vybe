// vybe-test: csharp/csharp_task_combinators/when_any_three_tasks_first_completed_count
// origin: languages/csharp/tests/csharp/test_csharp_task_combinators.rs

using static __Harness;

async System.Threading.Tasks.Task<int> A() { return 10; }
async System.Threading.Tasks.Task<int> B() { return 20; }
async System.Threading.Tasks.Task<int> C() { return 30; }
async System.Threading.Tasks.Task Run() {
    var winner = await System.Threading.Tasks.Task.WhenAny(A(), B(), C());
    __P((winner.Result).ToString());
}
Run().Wait();
__Check("10");

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

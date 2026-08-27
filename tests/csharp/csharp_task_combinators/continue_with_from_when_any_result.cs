// vybe-test: csharp/csharp_task_combinators/continue_with_from_when_any_result
// origin: languages/csharp/tests/csharp/test_csharp_task_combinators.rs

using static __Harness;

async System.Threading.Tasks.Task<int> A() { return 4; }
async System.Threading.Tasks.Task<int> B() { return 8; }
async System.Threading.Tasks.Task Run() {
    int count = 0;
    var winner = await System.Threading.Tasks.Task.WhenAny(A(), B());
    await winner.ContinueWith(t => count = t.Result + 1);
    __P((count).ToString());
}
Run().Wait();
__Check("5");

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

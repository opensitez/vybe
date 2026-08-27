// vybe-test: csharp/csharp_task_combinators/continue_with_three_link_chain_count
// origin: languages/csharp/tests/csharp/test_csharp_task_combinators.rs

using static __Harness;

async System.Threading.Tasks.Task Run() {
    int count = 0;
    await System.Threading.Tasks.Task.FromResult(1)
        .ContinueWith(t => t.Result + 1)
        .ContinueWith(t => t.Result + 1)
        .ContinueWith(t => count = t.Result + 1);
    __P((count).ToString());
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

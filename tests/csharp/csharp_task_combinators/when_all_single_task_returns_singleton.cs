// vybe-test: csharp/csharp_task_combinators/when_all_single_task_returns_singleton
// origin: languages/csharp/tests/csharp/test_csharp_task_combinators.rs

using static __Harness;

async System.Threading.Tasks.Task<int> Solo() { return 11; }
async System.Threading.Tasks.Task Run() {
    var results = await System.Threading.Tasks.Task.WhenAll(Solo());
    __P((results[0]).ToString());
}
Run().Wait();
__Check("11");

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

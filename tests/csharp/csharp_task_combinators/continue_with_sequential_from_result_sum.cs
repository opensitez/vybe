// vybe-test: csharp/csharp_task_combinators/continue_with_sequential_from_result_sum
// origin: languages/csharp/tests/csharp/test_csharp_task_combinators.rs

using static __Harness;

async System.Threading.Tasks.Task Run() {
    int count = 0;
    await System.Threading.Tasks.Task.FromResult(3)
        .ContinueWith(t => count += t.Result);
    await System.Threading.Tasks.Task.FromResult(4)
        .ContinueWith(t => count += t.Result);
    __P((count).ToString());
}
Run().Wait();
__Check("7");

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

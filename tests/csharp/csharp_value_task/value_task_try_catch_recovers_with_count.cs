// vybe-test: csharp/csharp_value_task/value_task_try_catch_recovers_with_count
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

using static __Harness;

async System.Threading.Tasks.ValueTask<int> Risky(bool fail) {
    if (fail) throw new System.Exception("no");
    return 4;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    try { count = await Risky(true); }
    catch (System.Exception) { count = 2; }
    __P((count).ToString());
}
Run().Wait();
__Check("2");

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

// vybe-test: csharp/csharp_value_task/value_task_array_length_after_await
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

using static __Harness;

async System.Threading.Tasks.ValueTask<int[]> Get() {
    return new int[] { 1, 2, 3, 4, 5 };
}
async System.Threading.Tasks.Task Run() {
    int[] arr = await Get();
    __P((arr.Length).ToString());
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

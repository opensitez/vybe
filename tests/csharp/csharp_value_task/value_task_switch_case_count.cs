// vybe-test: csharp/csharp_value_task/value_task_switch_case_count
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

using static __Harness;

async System.Threading.Tasks.ValueTask<int> Code() { return 2; }
async System.Threading.Tasks.Task Run() {
    int c = await Code();
    int count = 0;
    switch (c) {
        case 1: count = 10; break;
        case 2: count = 20; break;
        default: count = 0; break;
    }
    __P((count).ToString());
}
Run().Wait();
__Check("20");

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

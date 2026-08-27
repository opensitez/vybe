// vybe-test: csharp/csharp_task_combinators/task_run_three_parallel_workers_sum
// origin: languages/csharp/tests/csharp/test_csharp_task_combinators.rs

using static __Harness;

async System.Threading.Tasks.Task Run() {
    var a = System.Threading.Tasks.Task.Run(() => 1);
    var b = System.Threading.Tasks.Task.Run(() => 2);
    var c = System.Threading.Tasks.Task.Run(() => 3);
    __P((a.Result + b.Result + c.Result).ToString());
}
Run().Wait();
__Check("6");

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

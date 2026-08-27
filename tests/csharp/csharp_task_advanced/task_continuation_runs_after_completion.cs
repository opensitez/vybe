// vybe-test: csharp/csharp_task_advanced/task_continuation_runs_after_completion
// origin: languages/csharp/tests/csharp/test_csharp_task_advanced.rs

using static __Harness;

int result=0;
System.Threading.Tasks.Task.Run(()=>7)
    .ContinueWith(t=>result=t.Result*2)
    .Wait();
__P((result).ToString());
__Check("14");

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

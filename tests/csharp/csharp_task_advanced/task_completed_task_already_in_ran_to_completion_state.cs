// vybe-test: csharp/csharp_task_advanced/task_completed_task_already_in_ran_to_completion_state
// origin: languages/csharp/tests/csharp/test_csharp_task_advanced.rs

using static __Harness;

var t=System.Threading.Tasks.Task.CompletedTask;
__P((t.IsCompleted).ToString());
__Check("True");

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

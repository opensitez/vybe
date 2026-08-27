// vybe-test: csharp/csharp_task_advanced/when_any_returns_first_completed_task
// origin: languages/csharp/tests/csharp/test_csharp_task_advanced.rs

using static __Harness;

async System.Threading.Tasks.Task<int> Fast()=>await System.Threading.Tasks.Task.FromResult(1);
async System.Threading.Tasks.Task<int> Slow(){await System.Threading.Tasks.Task.Delay(1000);return 2;}
var winner=await System.Threading.Tasks.Task.WhenAny(Fast(),Slow());
__P((winner.Result).ToString());
__Check("1");

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

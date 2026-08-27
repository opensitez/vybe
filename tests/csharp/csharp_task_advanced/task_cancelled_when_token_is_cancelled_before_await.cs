// vybe-test: csharp/csharp_task_advanced/task_cancelled_when_token_is_cancelled_before_await
// origin: languages/csharp/tests/csharp/test_csharp_task_advanced.rs

using static __Harness;

var cts=new System.Threading.CancellationTokenSource();
cts.Cancel();
string result="ok";
try{System.Threading.Tasks.Task.Delay(1000,cts.Token).Wait();}
catch(System.AggregateException){result="cancelled";}
__P((result).ToString());
__Check("cancelled");

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

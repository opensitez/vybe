// vybe-test: csharp/csharp_task_advanced/task_from_exception_rethrows_on_await
// origin: languages/csharp/tests/csharp/test_csharp_task_advanced.rs

using static __Harness;

string msg="";
var t=System.Threading.Tasks.Task.FromException(new System.Exception("boom"));
try{t.Wait();}
catch(System.AggregateException ae){msg=ae.InnerException.Message;}
__P((msg).ToString());
__Check("boom");

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

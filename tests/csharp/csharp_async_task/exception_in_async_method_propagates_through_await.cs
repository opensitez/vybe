// vybe-test: csharp/csharp_async_task/exception_in_async_method_propagates_through_await
// origin: languages/csharp/tests/csharp/test_csharp_async_task.rs

using static __Harness;

async System.Threading.Tasks.Task Fail() {
    await System.Threading.Tasks.Task.Yield();
    throw new System.Exception("async fail");
}
string msg = "";
try { Fail().Wait(); }
catch (System.AggregateException ae) { msg = ae.InnerException.Message; }
__P((msg).ToString());
__Check("async fail");

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

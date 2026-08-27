// vybe-test: csharp/csharp_async_advanced/configure_await_false_does_not_resume_on_original_context
// origin: languages/csharp/tests/csharp/test_csharp_async_advanced.rs

using static __Harness;

async System.Threading.Tasks.Task<int> Compute(){
    await System.Threading.Tasks.Task.Delay(1).ConfigureAwait(false);
    return 42;
}
__P((await Compute()).ToString());
__Check("42");

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

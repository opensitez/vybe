// vybe-test: csharp/csharp_async_await_flow/async_void_local_function_can_be_awaited_from_enclosing_async_method
// origin: languages/csharp/tests/csharp/test_csharp_async_await_flow.rs

using static __Harness;
using System.Threading.Tasks;

async Task Run() {
    async Task<int> Compute() { return await Task.FromResult(6); }
    __P((await Compute()).ToString());
}
Run().GetAwaiter().GetResult();
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

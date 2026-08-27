// vybe-test: csharp/csharp_async_await_flow/await_on_task_from_result_returns_completed_value_to_caller
// origin: languages/csharp/tests/csharp/test_csharp_async_await_flow.rs

using static __Harness;
using System.Threading.Tasks;

async Task<int> Load() { return await Task.FromResult(9); }
async Task Run() { __P((await Load()).ToString()); }
Run().GetAwaiter().GetResult();
__Check("9");

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

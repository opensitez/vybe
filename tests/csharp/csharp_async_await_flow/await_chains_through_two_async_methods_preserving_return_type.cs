// vybe-test: csharp/csharp_async_await_flow/await_chains_through_two_async_methods_preserving_return_type
// origin: languages/csharp/tests/csharp/test_csharp_async_await_flow.rs

using static __Harness;
using System.Threading.Tasks;

async Task<int> Inner() { return await Task.FromResult(4); }
async Task<int> Outer() { return await Inner() + 1; }
__P((Outer().GetAwaiter().GetResult()).ToString());
__Check("5");

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

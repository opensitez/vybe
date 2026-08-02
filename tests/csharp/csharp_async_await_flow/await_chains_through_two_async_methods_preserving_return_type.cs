// vybe-test: csharp/csharp_async_await_flow/await_chains_through_two_async_methods_preserving_return_type
// origin: languages/csharp/tests/csharp/test_csharp_async_await_flow.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Threading.Tasks;
async Task<int> Inner() { return await Task.FromResult(4); }
async Task<int> Outer() { return await Inner() + 1; }
__Check((Outer().GetAwaiter().GetResult()).ToString(), "5");

// vybe-test: csharp/csharp_async_await_flow/async_void_local_function_can_be_awaited_from_enclosing_async_method
// origin: languages/csharp/tests/csharp/test_csharp_async_await_flow.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Threading.Tasks;
async Task Run() {
    async Task<int> Compute() { return await Task.FromResult(6); }
    __Check((await Compute()).ToString(), "6");
}
Run().GetAwaiter().GetResult();

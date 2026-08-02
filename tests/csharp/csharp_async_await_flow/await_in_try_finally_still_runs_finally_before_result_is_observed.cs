// vybe-test: csharp/csharp_async_await_flow/await_in_try_finally_still_runs_finally_before_result_is_observed
// origin: languages/csharp/tests/csharp/test_csharp_async_await_flow.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Threading.Tasks;
async Task<int> Pick() {
    try {
        return await Task.FromResult(2);
    } finally {
        __Check(("cleanup").ToString(), "cleanup");
    }
}
__Check((Pick().GetAwaiter().GetResult()).ToString(), "2");

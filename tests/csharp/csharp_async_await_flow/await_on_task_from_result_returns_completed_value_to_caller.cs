// vybe-test: csharp/csharp_async_await_flow/await_on_task_from_result_returns_completed_value_to_caller
// origin: languages/csharp/tests/csharp/test_csharp_async_await_flow.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Threading.Tasks;
async Task<int> Load() { return await Task.FromResult(9); }
async Task Run() { __Check((await Load()).ToString(), "9"); }
Run().GetAwaiter().GetResult();

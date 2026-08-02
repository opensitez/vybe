// vybe-test: csharp/csharp_async_await_flow/task_run_offloads_work_and_returns_result_to_awaiter
// origin: languages/csharp/tests/csharp/test_csharp_async_await_flow.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Threading.Tasks;
async Task<int> Run() {
    return await Task.Run(() => 11);
}
__Check((Run().GetAwaiter().GetResult()).ToString(), "11");

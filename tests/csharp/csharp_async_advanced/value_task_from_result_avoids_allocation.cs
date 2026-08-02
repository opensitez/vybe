// vybe-test: csharp/csharp_async_advanced/value_task_from_result_avoids_allocation
// origin: languages/csharp/tests/csharp/test_csharp_async_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.ValueTask<int> GetValueAsync()=>42;
int v=await GetValueAsync();
__Check((v).ToString(), "42");

// vybe-test: csharp/csharp_async_advanced/when_all_awaits_all_tasks_and_returns_results
// origin: languages/csharp/tests/csharp/test_csharp_async_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.Task<int> N(int v){
    await System.Threading.Tasks.Task.Delay(0);return v;
}
int[] results=await System.Threading.Tasks.Task.WhenAll(N(1),N(2),N(3));
__Check((results.Sum()).ToString(), "6");

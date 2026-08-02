// vybe-test: csharp/csharp_task_combinators/when_all_with_delay_zero_sum
// origin: languages/csharp/tests/csharp/test_csharp_task_combinators.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.Task<int> N(int v) {
    await System.Threading.Tasks.Task.Delay(0);
    return v;
}
async System.Threading.Tasks.Task Run() {
    var results = await System.Threading.Tasks.Task.WhenAll(N(2), N(3));
    __Check((results[0] + results[1]).ToString(), "5");
}
Run().Wait();

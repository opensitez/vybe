// vybe-test: csharp/csharp_task_combinators/when_all_empty_array_length_zero
// origin: languages/csharp/tests/csharp/test_csharp_task_combinators.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.Task Run() {
    var results = await System.Threading.Tasks.Task.WhenAll(new System.Threading.Tasks.Task<int>[0]);
    __Check((results.Length).ToString(), "0");
}
Run().Wait();

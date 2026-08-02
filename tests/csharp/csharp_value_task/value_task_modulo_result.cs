// vybe-test: csharp/csharp_value_task/value_task_modulo_result
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.ValueTask<int> Dividend() { return 17; }
async System.Threading.Tasks.ValueTask<int> Divisor() { return 5; }
async System.Threading.Tasks.Task Run() {
    __Check((await Dividend() % await Divisor()).ToString(), "2");
}
Run().Wait();

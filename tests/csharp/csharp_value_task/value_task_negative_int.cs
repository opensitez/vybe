// vybe-test: csharp/csharp_value_task/value_task_negative_int
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.ValueTask<int> Neg() { return -15; }
async System.Threading.Tasks.Task Run() {
    __Check((await Neg()).ToString(), "-15");
}
Run().Wait();

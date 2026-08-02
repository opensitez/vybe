// vybe-test: csharp/csharp_value_task/value_task_double_result
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.ValueTask<double> Get() { return 3.5; }
async System.Threading.Tasks.Task Run() {
    double v = await Get();
    __Check(((int)v).ToString(), "3");
}
Run().Wait();

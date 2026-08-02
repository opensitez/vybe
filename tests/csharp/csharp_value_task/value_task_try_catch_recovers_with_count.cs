// vybe-test: csharp/csharp_value_task/value_task_try_catch_recovers_with_count
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.ValueTask<int> Risky(bool fail) {
    if (fail) throw new System.Exception("no");
    return 4;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    try { count = await Risky(true); }
    catch (System.Exception) { count = 2; }
    __Check((count).ToString(), "2");
}
Run().Wait();

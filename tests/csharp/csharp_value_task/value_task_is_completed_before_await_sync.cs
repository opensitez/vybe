// vybe-test: csharp/csharp_value_task/value_task_is_completed_before_await_sync
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.ValueTask<int> Sync() { return 11; }
async System.Threading.Tasks.Task Run() {
    var vt = Sync();
    int count = vt.IsCompleted ? 1 : 0;
    __Check((count).ToString(), "1");
}
Run().Wait();

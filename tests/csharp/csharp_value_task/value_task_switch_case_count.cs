// vybe-test: csharp/csharp_value_task/value_task_switch_case_count
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.ValueTask<int> Code() { return 2; }
async System.Threading.Tasks.Task Run() {
    int c = await Code();
    int count = 0;
    switch (c) {
        case 1: count = 10; break;
        case 2: count = 20; break;
        default: count = 0; break;
    }
    __Check((count).ToString(), "20");
}
Run().Wait();

// vybe-test: csharp/csharp_value_task/value_task_if_branch_false_path_count
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.ValueTask<int> Pick(bool flag) {
    if (flag) return 100;
    return 7;
}
async System.Threading.Tasks.Task Run() {
    __Check((await Pick(false)).ToString(), "7");
}
Run().Wait();

// vybe-test: csharp/csharp_value_task/value_task_list_count_property
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.ValueTask<System.Collections.Generic.List<int>> Get() {
    return new System.Collections.Generic.List<int> { 1, 2, 3 };
}
async System.Threading.Tasks.Task Run() {
    var list = await Get();
    __Check((list.Count).ToString(), "3");
}
Run().Wait();

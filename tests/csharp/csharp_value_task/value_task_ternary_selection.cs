// vybe-test: csharp/csharp_value_task/value_task_ternary_selection
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.ValueTask<int> Choose(int a, int b, bool first) {
    return first ? a : b;
}
async System.Threading.Tasks.Task Run() {
    __Check((await Choose(3, 8, false)).ToString(), "8");
}
Run().Wait();

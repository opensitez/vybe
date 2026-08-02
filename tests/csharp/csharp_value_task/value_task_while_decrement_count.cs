// vybe-test: csharp/csharp_value_task/value_task_while_decrement_count
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

async System.Threading.Tasks.ValueTask<int> Step() { return 1; }
async System.Threading.Tasks.Task Run() {
    int n = 4;
    int count = 0;
    while (n > 0) { count += await Step(); n--; }
    Console.WriteLine(count);
}
Run().Wait();

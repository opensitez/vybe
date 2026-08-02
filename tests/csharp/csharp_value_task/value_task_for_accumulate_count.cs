// vybe-test: csharp/csharp_value_task/value_task_for_accumulate_count
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

async System.Threading.Tasks.ValueTask<int> One() { return 1; }
async System.Threading.Tasks.Task Run() {
    int count = 0;
    for (int i = 0; i < 8; i++) count += await One();
    Console.WriteLine(count);
}
Run().Wait();

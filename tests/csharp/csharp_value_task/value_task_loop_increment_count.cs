// vybe-test: csharp/csharp_value_task/value_task_loop_increment_count
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

async System.Threading.Tasks.ValueTask<int> Step() { return 1; }
async System.Threading.Tasks.Task Run() {
    int count = 0;
    for (int i = 0; i < 5; i++) count += await Step();
    Console.WriteLine(count);
}
Run().Wait();

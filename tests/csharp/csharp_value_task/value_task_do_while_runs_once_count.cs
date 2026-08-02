// vybe-test: csharp/csharp_value_task/value_task_do_while_runs_once_count
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

async System.Threading.Tasks.ValueTask<int> Step() { return 1; }
async System.Threading.Tasks.Task Run() {
    int count = 0;
    do { count += await Step(); } while (false);
    Console.WriteLine(count);
}
Run().Wait();

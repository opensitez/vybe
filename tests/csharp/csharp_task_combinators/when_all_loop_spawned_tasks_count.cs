// vybe-test: csharp/csharp_task_combinators/when_all_loop_spawned_tasks_count
// origin: languages/csharp/tests/csharp/test_csharp_task_combinators.rs

async System.Threading.Tasks.Task<int> N(int v) { return v; }
async System.Threading.Tasks.Task Run() {
    var tasks = new System.Threading.Tasks.Task<int>[3];
    for (int i = 0; i < 3; i++) tasks[i] = N(i + 1);
    var results = await System.Threading.Tasks.Task.WhenAll(tasks);
    int count = 0;
    foreach (var x in results) count += x;
    Console.WriteLine(count);
}
Run().Wait();

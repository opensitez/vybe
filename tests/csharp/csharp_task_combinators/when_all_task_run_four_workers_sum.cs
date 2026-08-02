// vybe-test: csharp/csharp_task_combinators/when_all_task_run_four_workers_sum
// origin: languages/csharp/tests/csharp/test_csharp_task_combinators.rs

async System.Threading.Tasks.Task Run() {
    var results = await System.Threading.Tasks.Task.WhenAll(
        System.Threading.Tasks.Task.Run(() => 1),
        System.Threading.Tasks.Task.Run(() => 2),
        System.Threading.Tasks.Task.Run(() => 3),
        System.Threading.Tasks.Task.Run(() => 4)
    );
    int count = 0;
    foreach (var x in results) count += x;
    Console.WriteLine(count);
}
Run().Wait();

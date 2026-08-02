// vybe-test: csharp/csharp_task_combinators/when_all_four_tasks_product_count
// origin: languages/csharp/tests/csharp/test_csharp_task_combinators.rs

async System.Threading.Tasks.Task<int> N(int v) { return v; }
async System.Threading.Tasks.Task Run() {
    var results = await System.Threading.Tasks.Task.WhenAll(N(2), N(3), N(4), N(5));
    int count = 1;
    foreach (var x in results) count *= x;
    Console.WriteLine(count);
}
Run().Wait();
